use portable_pty::{native_pty_system, Child, CommandBuilder, MasterPty, PtySize};
use serde::{Deserialize, Serialize};
use std::{
    ffi::c_void,
    collections::HashMap,
    cell::Cell,
    io::{BufReader, Read, Write},
    path::{Path, PathBuf},
    process::Command,
    sync::{Arc, Condvar, Mutex},
    thread,
    time::{Duration, SystemTime, UNIX_EPOCH},
};
use tauri::{AppHandle, Emitter, Manager};
use std::mem::ManuallyDrop;
use windows::{
    core::Interface,
    Win32::{
        Foundation::{
            DRAGDROP_S_CANCEL, DRAGDROP_S_DROP, DRAGDROP_S_USEDEFAULTCURSORS, DV_E_FORMATETC,
            E_FAIL, E_NOTIMPL, E_UNEXPECTED, HGLOBAL, S_FALSE, S_OK,
        },
        System::{
            Com::{
                CoInitializeEx, CoUninitialize, IDataObject, IDataObject_Impl, FORMATETC,
                IEnumFORMATETC, IEnumFORMATETC_Impl, STGMEDIUM, TYMED_HGLOBAL, TYMED_ISTREAM,
                ISequentialStream_Impl, IStream, IStream_Impl, LOCKTYPE, STATFLAG, STATSTG, STGC,
                STREAM_SEEK, STGTY_STREAM,
                COINIT_APARTMENTTHREADED, DATADIR_GET,
                StructuredStorage::CreateStreamOnHGlobal,
            },
            DataExchange::RegisterClipboardFormatW,
            Memory::{GlobalAlloc, GlobalLock, GlobalUnlock, GMEM_MOVEABLE, GMEM_ZEROINIT},
            Ole::{
                CF_HDROP, DoDragDrop, DROPEFFECT, DROPEFFECT_COPY, IDropSource, IDropSource_Impl,
                OleInitialize, OleUninitialize,
            },
            SystemServices::MODIFIERKEYS_FLAGS,
        },
        UI::Shell::{
            Common, SHCreateDataObject, SHParseDisplayName, DROPFILES, FILEDESCRIPTORW,
            FILEGROUPDESCRIPTORW, FD_ATTRIBUTES, FD_FILESIZE, FD_PROGRESSUI,
        },
    },
};
use windows_core::{implement, w, BOOL, HRESULT, Result as WinResult};

#[cfg(target_os = "windows")]
unsafe extern "C" {
    fn mmshell_get_file_size_utf8(path_utf8: *const i8) -> u64;
    fn mmshell_read_file_chunk_utf8(
        path_utf8: *const i8,
        offset: u64,
        max_len: u32,
        out_buf: *mut u8,
        out_read: *mut u32,
    ) -> i32;
    fn mmshell_start_virtual_drag_from_file_utf8(
        local_path_utf8: *const i8,
        display_name_utf8: *const i8,
        out_effect: *mut u32,
    ) -> i32;
    fn mmshell_start_hdrop_drag_from_file_utf8(
        local_path_utf8: *const i8,
        out_effect: *mut u32,
    ) -> i32;
    fn mmshell_start_virtual_drag_streaming_utf8(
        local_path_utf8: *const i8,
        display_name_utf8: *const i8,
        done_marker_utf8: *const i8,
        err_marker_utf8: *const i8,
        wait_ms: u32,
        out_effect: *mut u32,
    ) -> i32;
    fn mmshell_detect_drop_target_utf8(
        out_path_utf8: *mut i8,
        out_path_capacity: u32,
        out_effect: *mut u32,
    ) -> i32;
}

#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
struct SshOutputEvent {
    session_id: String,
    data: String,
    stream: String,
}

struct ActiveSession {
    writer: Arc<Mutex<Box<dyn Write + Send>>>,
    master: Arc<Mutex<Box<dyn MasterPty + Send>>>,
    child: Arc<Mutex<Box<dyn Child + Send + Sync>>>,
    command_lock: Arc<Mutex<()>>,
}

#[derive(Default)]
struct SshSessionStore {
    sessions: Mutex<HashMap<String, ActiveSession>>,
}

#[derive(Default)]
struct SftpSessionStore {
    sessions: Mutex<HashMap<String, ActiveSession>>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct ConnectPayload {
    address: String,
    /// 由前端传入；SSH/SFTP 在 PTY 中交互输入密码，此处仅反序列化以保持 API 一致。
    #[allow(dead_code)]
    password: String,
}

#[derive(Debug)]
struct ParsedAddress {
    user: String,
    host: String,
    port: u16,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct NativeDragOutPayload {
    remote_path: String,
    is_dir: bool,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct NativeDragOutResult {
    ok: bool,
    cache_path: String,
    message: String,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct CopyCachedPayload {
    source_path: String,
    target_dir: String,
    is_dir: bool,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct NativeDragDropResult {
    effect: u32,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct NativeVirtualDragPayload {
    source_path: String,
    display_name: String,
    sftp_session_id: Option<String>,
    remote_name: Option<String>,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct NativeSftpHdropPayload {
    sftp_session_id: String,
    remote_path: String,
    display_name: String,
    is_dir: bool,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct NativePlaceholderDragPayload {
    display_name: String,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct NativeCachedVirtualDragPayload {
    path: String,
    display_name: String,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct NativeSftpVirtualStreamingPayload {
    sftp_session_id: String,
    remote_path: String,
    display_name: String,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct NativeSftpDropPickPayload {
    sftp_session_id: String,
    remote_path: String,
    display_name: String,
    is_dir: bool,
}


/// 解析 user@host:port 地址字符串。 / Parse address string in user@host:port format.
fn parse_address(raw: &str) -> Result<ParsedAddress, String> {
    // Accept address format like: user@host:22 (port is optional).
    let trimmed = raw.trim();
    let (user_part, host_port_part) = trimmed
        .split_once('@')
        .ok_or_else(|| "地址格式必须是 user@host:port".to_string())?;
    if user_part.trim().is_empty() {
        return Err("用户名不能为空".to_string());
    }
    let (host, port) = match host_port_part.rsplit_once(':') {
        Some((h, p)) if !p.trim().is_empty() => {
            let parsed_port = p
                .trim()
                .parse::<u16>()
                .map_err(|_| "端口必须是 1-65535 的数字".to_string())?;
            (h.trim().to_string(), parsed_port)
        }
        _ => (host_port_part.trim().to_string(), 22),
    };
    if host.is_empty() {
        return Err("主机不能为空".to_string());
    }
    if port == 0 {
        return Err("端口必须大于 0".to_string());
    }
    Ok(ParsedAddress {
        user: user_part.trim().to_string(),
        host,
        port,
    })
}

/// 生成 SSH 会话唯一 ID。 / Generate unique SSH session id.
fn new_session_id() -> String {
    let ts = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_millis())
        .unwrap_or_default();
    format!("ssh-{ts}")
}

/// 生成 SFTP 会话唯一 ID。 / Generate unique SFTP session id.
fn new_sftp_session_id() -> String {
    let ts = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_millis())
        .unwrap_or_default();
    format!("sftp-{ts}")
}

/// 建立 SSH PTY 会话并启动输出转发。 / Connect SSH PTY session and forward output events.
#[tauri::command]
fn connect_ssh(app: AppHandle, payload: ConnectPayload) -> Result<String, String> {
    let parsed = parse_address(&payload.address)?;
    let session_id = new_session_id();

    let pty_system = native_pty_system();
    // Allocate a pseudo terminal so ssh.exe behaves like an interactive terminal.
    let pty_pair = pty_system
        .openpty(PtySize {
            rows: 40,
            cols: 120,
            pixel_width: 0,
            pixel_height: 0,
        })
        .map_err(|err| format!("创建 PTY 失败: {err}"))?;

    let mut cmd = CommandBuilder::new("ssh.exe");
    // Improve compatibility with embedded SSH servers (e.g. Luckfox / Dropbear).
    // - accept-new: avoid blocking on first host key confirmation prompt.
    // - ssh-rsa options: allow legacy algorithms still used by some devices.
    cmd.arg("-tt");
    cmd.arg("-o");
    cmd.arg("StrictHostKeyChecking=accept-new");
    cmd.arg("-o");
    cmd.arg("HostKeyAlgorithms=+ssh-rsa");
    cmd.arg("-o");
    cmd.arg("PubkeyAcceptedAlgorithms=+ssh-rsa");
    cmd.arg("-p");
    cmd.arg(parsed.port.to_string());
    cmd.arg(format!("{}@{}", parsed.user, parsed.host));

    let child = pty_pair
        .slave
        .spawn_command(cmd)
        .map_err(|err| format!("启动 ssh.exe 失败: {err}"))?;
    drop(pty_pair.slave);

    let reader = pty_pair
        .master
        .try_clone_reader()
        .map_err(|err| format!("创建 PTY reader 失败: {err}"))?;
    let writer = pty_pair
        .master
        .take_writer()
        .map_err(|err| format!("创建 PTY writer 失败: {err}"))?;

    let active = ActiveSession {
        writer: Arc::new(Mutex::new(writer)),
        master: Arc::new(Mutex::new(pty_pair.master)),
        child: Arc::new(Mutex::new(child)),
        command_lock: Arc::new(Mutex::new(())),
    };

    {
        // Store all session handles so other commands can write/resize/kill by session id.
        let store = app.state::<SshSessionStore>();
        let mut lock = store
            .sessions
            .lock()
            .map_err(|_| "session store lock poisoned".to_string())?;
        lock.insert(session_id.clone(), active);
    }

    let app_reader = app.clone();
    let sid_reader = session_id.clone();
    thread::spawn(move || {
        // Forward PTY output to frontend xterm via Tauri events.
        let mut reader = BufReader::new(reader);
        let mut buf = [0_u8; 4096];
        loop {
            match reader.read(&mut buf) {
                Ok(0) => break,
                Ok(n) => {
                    let text = String::from_utf8_lossy(&buf[..n]).to_string();
                    let _ = app_reader.emit(
                        "ssh-output",
                        SshOutputEvent {
                            session_id: sid_reader.clone(),
                            data: text,
                            stream: "stdout".to_string(),
                        },
                    );
                }
                Err(_) => break,
            }
        }
    });

    Ok(session_id)
}

/// 向指定 SSH 会话写入输入。 / Send input to target SSH session.
#[tauri::command]
fn send_ssh_input(app: AppHandle, session_id: String, input: String) -> Result<(), String> {
    let store = app.state::<SshSessionStore>();
    let lock = store
        .sessions
        .lock()
        .map_err(|_| "session store lock poisoned".to_string())?;
    let active = lock
        .get(&session_id)
        .ok_or_else(|| "ssh session not found".to_string())?;
    let mut writer = active
        .writer
        .lock()
        .map_err(|_| "pty writer lock poisoned".to_string())?;
    writer
        .write_all(input.as_bytes())
        .map_err(|err| format!("写入 SSH 输入失败: {err}"))?;
    writer
        .flush()
        .map_err(|err| format!("刷新 SSH 输入失败: {err}"))?;
    Ok(())
}

/// 调整 SSH PTY 终端尺寸。 / Resize SSH PTY terminal dimensions.
#[tauri::command]
fn resize_ssh(app: AppHandle, session_id: String, cols: u16, rows: u16) -> Result<(), String> {
    // Protect against invalid tiny sizes from intermediate window layouts.
    let safe_cols = cols.max(20);
    let safe_rows = rows.max(10);
    let store = app.state::<SshSessionStore>();
    let lock = store
        .sessions
        .lock()
        .map_err(|_| "session store lock poisoned".to_string())?;
    let active = lock
        .get(&session_id)
        .ok_or_else(|| "ssh session not found".to_string())?;
    let master = active
        .master
        .lock()
        .map_err(|_| "pty master lock poisoned".to_string())?;
    master
        .resize(PtySize {
            rows: safe_rows,
            cols: safe_cols,
            pixel_width: 0,
            pixel_height: 0,
        })
        .map_err(|err| format!("调整终端大小失败: {err}"))?;
    Ok(())
}

/// 断开并清理 SSH 会话资源。 / Disconnect and clean up SSH session resources.
#[tauri::command]
fn disconnect_ssh(app: AppHandle, session_id: String) -> Result<(), String> {
    let store = app.state::<SshSessionStore>();
    let mut lock = store
        .sessions
        .lock()
        .map_err(|_| "session store lock poisoned".to_string())?;
    if let Some(active) = lock.remove(&session_id) {
        if let Ok(mut child) = active.child.lock() {
            let _ = child.kill();
        }
    }
    Ok(())
}

/// 建立 SFTP PTY 会话并转发输出。 / Connect SFTP PTY session and forward output.
#[tauri::command]
fn connect_sftp(app: AppHandle, payload: ConnectPayload) -> Result<String, String> {
    let parsed = parse_address(&payload.address)?;
    let session_id = new_sftp_session_id();

    let pty_system = native_pty_system();
    let pty_pair = pty_system
        .openpty(PtySize {
            rows: 24,
            cols: 120,
            pixel_width: 0,
            pixel_height: 0,
        })
        .map_err(|err| format!("创建 SFTP PTY 失败: {err}"))?;

    let mut cmd = CommandBuilder::new("sftp.exe");
    // Keep SFTP and SSH compatibility options aligned.
    cmd.arg("-o");
    cmd.arg("StrictHostKeyChecking=accept-new");
    cmd.arg("-o");
    cmd.arg("HostKeyAlgorithms=+ssh-rsa");
    cmd.arg("-o");
    cmd.arg("PubkeyAcceptedAlgorithms=+ssh-rsa");
    cmd.arg("-P");
    cmd.arg(parsed.port.to_string());
    cmd.arg(format!("{}@{}", parsed.user, parsed.host));

    let child = pty_pair
        .slave
        .spawn_command(cmd)
        .map_err(|err| format!("启动 sftp.exe 失败: {err}"))?;
    drop(pty_pair.slave);

    let reader = pty_pair
        .master
        .try_clone_reader()
        .map_err(|err| format!("创建 SFTP reader 失败: {err}"))?;
    let writer = pty_pair
        .master
        .take_writer()
        .map_err(|err| format!("创建 SFTP writer 失败: {err}"))?;

    let active = ActiveSession {
        writer: Arc::new(Mutex::new(writer)),
        master: Arc::new(Mutex::new(pty_pair.master)),
        child: Arc::new(Mutex::new(child)),
        command_lock: Arc::new(Mutex::new(())),
    };

    {
        let store = app.state::<SftpSessionStore>();
        let mut lock = store
            .sessions
            .lock()
            .map_err(|_| "sftp session store lock poisoned".to_string())?;
        lock.insert(session_id.clone(), active);
    }

    let app_reader = app.clone();
    let sid_reader = session_id.clone();
    thread::spawn(move || {
        let mut reader = BufReader::new(reader);
        let mut buf = [0_u8; 4096];
        loop {
            match reader.read(&mut buf) {
                Ok(0) => break,
                Ok(n) => {
                    let text = String::from_utf8_lossy(&buf[..n]).to_string();
                    let _ = app_reader.emit(
                        "sftp-output",
                        SshOutputEvent {
                            session_id: sid_reader.clone(),
                            data: text,
                            stream: "stdout".to_string(),
                        },
                    );
                }
                Err(_) => break,
            }
        }
    });

    Ok(session_id)
}

/// 向 SFTP 会话发送命令输入。 / Send command input to SFTP session.
#[tauri::command]
fn send_sftp_input(app: AppHandle, session_id: String, input: String) -> Result<(), String> {
    let store = app.state::<SftpSessionStore>();
    let lock = store
        .sessions
        .lock()
        .map_err(|_| "sftp session store lock poisoned".to_string())?;
    let active = lock
        .get(&session_id)
        .ok_or_else(|| "sftp session not found".to_string())?;
    let _cmd_guard = active
        .command_lock
        .lock()
        .map_err(|_| "sftp command lock poisoned".to_string())?;
    let mut writer = active
        .writer
        .lock()
        .map_err(|_| "sftp writer lock poisoned".to_string())?;
    writer
        .write_all(input.as_bytes())
        .map_err(|err| format!("写入 SFTP 输入失败: {err}"))?;
    writer
        .flush()
        .map_err(|err| format!("刷新 SFTP 输入失败: {err}"))?;
    Ok(())
}

/// 断开并清理 SFTP 会话资源。 / Disconnect and clean up SFTP session resources.
#[tauri::command]
fn disconnect_sftp(app: AppHandle, session_id: String) -> Result<(), String> {
    let store = app.state::<SftpSessionStore>();
    let mut lock = store
        .sessions
        .lock()
        .map_err(|_| "sftp session store lock poisoned".to_string())?;
    if let Some(active) = lock.remove(&session_id) {
        if let Ok(mut child) = active.child.lock() {
            let _ = child.kill();
        }
    }
    Ok(())
}

/// 与前端 `App.tsx` 中 `MMSHELL_CONFIG_DIR` / `SESSIONS_CONFIG_PATH` 保持一致。
/// 初始化本地配置文件（不存在则创建）。 / Ensure local config file exists.
fn ensure_mmshell_config_file() {
    let dir = PathBuf::from(r"D:\MMShell0414");
    if let Err(err) = std::fs::create_dir_all(&dir) {
        eprintln!("[mmshell] 创建配置目录失败 {dir:?}: {err}");
        return;
    }
    let path = dir.join("mmshell_config.json");
    if path.exists() {
        return;
    }
    let initial = serde_json::json!({
        "version": 1,
        "sessions": [],
        "lastUpdated": ""
    });
    match serde_json::to_string_pretty(&initial) {
        Ok(content) => match std::fs::write(&path, content) {
            Ok(()) => eprintln!("[mmshell] 已创建会话配置文件: {}", path.display()),
            Err(err) => eprintln!("[mmshell] 写入配置文件失败 {}: {err}", path.display()),
        },
        Err(err) => eprintln!("[mmshell] 序列化默认配置失败: {err}"),
    }
}

/// 清洗文件名中的非法字符并限制长度。 / Sanitize file name by removing invalid characters.
fn sanitize_placeholder_name(raw: &str) -> String {
    let mut name = raw
        .chars()
        .map(|c| {
            if ['\\', '/', ':', '*', '?', '"', '<', '>', '|'].contains(&c) {
                '_'
            } else {
                c
            }
        })
        .collect::<String>();
    if name.trim().is_empty() {
        name = "drag-item".to_string();
    }
    if name.len() > 80 {
        name.truncate(80);
    }
    name
}

/// 在资源管理器中选中目标路径。 / Reveal target path in Windows Explorer.
fn open_in_explorer(path: &Path) {
    let _ = Command::new("explorer.exe")
        .arg("/select,")
        .arg(path)
        .spawn();
}

/// 按 CF_HDROP 规范构造全局内存数据块。 / Build HGLOBAL data block in CF_HDROP format.
fn build_hdrop_for_path(path: &str) -> Result<HGLOBAL, String> {
    let canonical = PathBuf::from(path)
        .canonicalize()
        .map_err(|e| format!("构造 CF_HDROP 前路径无效: {e}"))?;
    let mut absolute = canonical.to_string_lossy().replace('/', "\\");
    // 某些目标（尤其资源管理器）对 \\?\ 扩展前缀兼容不好，CF_HDROP 用普通绝对路径更稳。
    if let Some(stripped) = absolute.strip_prefix("\\\\?\\UNC\\") {
        absolute = format!("\\\\{}", stripped);
    } else if let Some(stripped) = absolute.strip_prefix("\\\\?\\") {
        absolute = stripped.to_string();
    }
    // CF_HDROP 文件列表必须是: path1\0path2\0\0，单文件也必须双 \0 结尾。
    let mut wide: Vec<u16> = absolute.encode_utf16().collect();
    wide.push(0); // end of first path
    wide.push(0); // end of list
    let header_size = std::mem::size_of::<DROPFILES>();
    let payload_size = wide.len() * std::mem::size_of::<u16>();
    let total_size = header_size + payload_size;

    let hglobal = unsafe { GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, total_size) }
        .map_err(|e| format!("GlobalAlloc 失败: {e}"))?;
    let locked = unsafe { GlobalLock(hglobal) } as *mut u8;
    if locked.is_null() {
        return Err("GlobalLock 失败".to_string());
    }
    unsafe {
        let dropfiles = locked as *mut DROPFILES;
        (*dropfiles).pFiles = header_size as u32;
        (*dropfiles).fNC = BOOL(0);
        (*dropfiles).fWide = BOOL(1);
        let dst = locked.add(header_size) as *mut u16;
        std::ptr::copy_nonoverlapping(wide.as_ptr(), dst, wide.len());
        let _ = GlobalUnlock(hglobal);
    }
    eprintln!(
        "[mmshell][hdrop] path={} header_size={} pFiles={} payload_u16_len={}",
        absolute,
        header_size,
        header_size,
        wide.len()
    );
    Ok(hglobal)
}

/// 获取 Preferred DropEffect 剪贴板格式 ID。 / Get clipboard format id for Preferred DropEffect.
fn cf_preferred_drop_effect() -> u16 {
    unsafe { RegisterClipboardFormatW(w!("Preferred DropEffect")) as u16 }
}

/// 获取 Performed DropEffect 剪贴板格式 ID。 / Get clipboard format id for Performed DropEffect.
fn cf_performed_drop_effect() -> u16 {
    unsafe { RegisterClipboardFormatW(w!("Performed DropEffect")) as u16 }
}

/// 获取 Paste Succeeded 剪贴板格式 ID。 / Get clipboard format id for Paste Succeeded.
fn cf_paste_succeeded() -> u16 {
    unsafe { RegisterClipboardFormatW(w!("Paste Succeeded")) as u16 }
}

/// 构造包含拖放效果值的 HGLOBAL。 / Build HGLOBAL containing a drop-effect u32.
fn build_drop_effect_hglobal(effect: u32) -> Result<HGLOBAL, String> {
    let hglobal = unsafe { GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, 4) }
        .map_err(|e| format!("GlobalAlloc(DropEffect) 失败: {e}"))?;
    let locked = unsafe { GlobalLock(hglobal) } as *mut u8;
    if locked.is_null() {
        return Err("GlobalLock(DropEffect) 失败".to_string());
    }
    unsafe {
        std::ptr::copy_nonoverlapping(effect.to_le_bytes().as_ptr(), locked, 4);
        let _ = GlobalUnlock(hglobal);
    }
    Ok(hglobal)
}

#[implement(IDataObject, IDropSource)]
struct NativeDragObject {
    path: String,
}

#[implement(IEnumFORMATETC)]
struct NativeFormatEtcEnum {
    index: Cell<usize>,
    items: Vec<FORMATETC>,
}

impl NativeFormatEtcEnum {
    /// 创建 FORMATETC 枚举器实例。 / Create a FORMATETC enumerator instance.
    fn new(items: Vec<FORMATETC>) -> Self {
        Self {
            index: Cell::new(0),
            items,
        }
    }
}

impl IEnumFORMATETC_Impl for NativeFormatEtcEnum_Impl {
    /// 返回下一批格式项。 / Return next batch of clipboard formats.
    fn Next(&self, celt: u32, rgelt: *mut FORMATETC, pceltfetched: *mut u32) -> HRESULT {
        if celt == 0 || rgelt.is_null() {
            return DV_E_FORMATETC;
        }
        let mut fetched = 0usize;
        let mut idx = self.index.get();
        while fetched < celt as usize && idx < self.items.len() {
            unsafe {
                *rgelt.add(fetched) = self.items[idx];
            }
            fetched += 1;
            idx += 1;
        }
        self.index.set(idx);
        if !pceltfetched.is_null() {
            unsafe { *pceltfetched = fetched as u32; }
        }
        if fetched == celt as usize {
            S_OK
        } else {
            S_FALSE
        }
    }

    /// 跳过指定数量的格式项。 / Skip a given number of formats.
    fn Skip(&self, celt: u32) -> windows_core::Result<()> {
        let idx = self.index.get().saturating_add(celt as usize);
        self.index.set(idx.min(self.items.len()));
        Ok(())
    }

    /// 重置枚举游标到起点。 / Reset enumeration cursor to start.
    fn Reset(&self) -> windows_core::Result<()> {
        self.index.set(0);
        Ok(())
    }

    /// 克隆当前枚举器状态。 / Clone current enumerator state.
    fn Clone(&self) -> windows_core::Result<IEnumFORMATETC> {
        let cloned = NativeFormatEtcEnum {
            index: Cell::new(self.index.get()),
            items: self.items.clone(),
        };
        Ok(cloned.into())
    }
}

impl IDataObject_Impl for NativeDragObject_Impl {
    /// 按请求格式提供拖拽数据。 / Provide drag payload for requested format.
    fn GetData(&self, pformatetcin: *const FORMATETC) -> WinResult<STGMEDIUM> {
        if pformatetcin.is_null() {
            return Err(windows_core::Error::new(E_UNEXPECTED, "FORMATETC is null"));
        }
        let fmt = unsafe { *pformatetcin };
        let mut medium = STGMEDIUM::default();
        medium.tymed = TYMED_HGLOBAL.0 as u32;
        if fmt.cfFormat == CF_HDROP.0 && fmt.tymed == TYMED_HGLOBAL.0 as u32 {
            let hglobal = build_hdrop_for_path(&self.path)
                .map_err(|e| windows_core::Error::new(E_UNEXPECTED, e.to_string()))?;
            medium.u.hGlobal = hglobal;
            medium.pUnkForRelease = ManuallyDrop::new(None);
            return Ok(medium);
        }
        if fmt.cfFormat == cf_preferred_drop_effect() && fmt.tymed == TYMED_HGLOBAL.0 as u32 {
            let hglobal = unsafe { GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, 4) }
                .map_err(|e| windows_core::Error::new(E_UNEXPECTED, format!("GlobalAlloc(PreferredDropEffect) 失败: {e}")))?;
            let locked = unsafe { GlobalLock(hglobal) } as *mut u8;
            if locked.is_null() {
                return Err(windows_core::Error::new(E_UNEXPECTED, "GlobalLock(PreferredDropEffect) 失败"));
            }
            unsafe {
                let effect_copy = DROPEFFECT_COPY.0 as u32;
                std::ptr::copy_nonoverlapping(effect_copy.to_le_bytes().as_ptr(), locked, 4);
                let _ = GlobalUnlock(hglobal);
            }
            medium.u.hGlobal = hglobal;
            medium.pUnkForRelease = ManuallyDrop::new(None);
            return Ok(medium);
        }
        Err(windows_core::Error::new(DV_E_FORMATETC, "unsupported format"))
    }

    /// 不支持由目标提供缓冲区的写入模式。 / GetDataHere path is not supported.
    fn GetDataHere(&self, _pformatetc: *const FORMATETC, _pmedium: *mut STGMEDIUM) -> WinResult<()> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
    /// 检查当前格式是否可被提供。 / Check whether the requested format is supported.
    fn QueryGetData(&self, pformatetc: *const FORMATETC) -> HRESULT {
        if pformatetc.is_null() {
            return DV_E_FORMATETC;
        }
        let fmt = unsafe { *pformatetc };
        if (fmt.cfFormat == CF_HDROP.0 || fmt.cfFormat == cf_preferred_drop_effect())
            && fmt.tymed == TYMED_HGLOBAL.0 as u32
        {
            S_OK
        } else {
            DV_E_FORMATETC
        }
    }
    /// 不提供规范化格式映射。 / Canonical format mapping is not implemented.
    fn GetCanonicalFormatEtc(&self, _a: *const FORMATETC, _b: *mut FORMATETC) -> HRESULT { E_NOTIMPL }
    /// 处理目标端回写格式（如 DropEffect）。 / Handle target-side write-back formats.
    fn SetData(&self, _a: *const FORMATETC, _b: *const STGMEDIUM, _c: BOOL) -> WinResult<()> {
        // Explorer 可能在拖放结束时回写这些格式；接受它们可提升 shell 兼容性。
        if _a.is_null() {
            return Err(windows_core::Error::from(DV_E_FORMATETC));
        }
        let fmt = unsafe { *_a };
        if fmt.cfFormat == cf_performed_drop_effect() {
            eprintln!("[mmshell][hdrop] SetData: PerformedDropEffect received");
            return Ok(());
        }
        if fmt.cfFormat == cf_paste_succeeded() {
            eprintln!("[mmshell][hdrop] SetData: PasteSucceeded received");
            return Ok(());
        }
        eprintln!(
            "[mmshell][hdrop] SetData: ignored format cf={} tymed={}",
            fmt.cfFormat,
            fmt.tymed
        );
        Err(windows_core::Error::from(E_NOTIMPL))
    }
    /// 枚举当前对象可提供的格式列表。 / Enumerate all formats this object can provide.
    fn EnumFormatEtc(&self, dwdirection: u32) -> WinResult<windows::Win32::System::Com::IEnumFORMATETC> {
        if dwdirection != DATADIR_GET.0 as u32 {
            return Err(windows_core::Error::from(DV_E_FORMATETC));
        }
        let fmts = vec![
            FORMATETC {
                cfFormat: CF_HDROP.0,
                ptd: std::ptr::null_mut(),
                dwAspect: 1, // DVASPECT_CONTENT
                lindex: -1,
                tymed: TYMED_HGLOBAL.0 as u32,
            },
            FORMATETC {
                cfFormat: cf_preferred_drop_effect(),
                ptd: std::ptr::null_mut(),
                dwAspect: 1,
                lindex: -1,
                tymed: TYMED_HGLOBAL.0 as u32,
            },
            FORMATETC {
                cfFormat: cf_performed_drop_effect(),
                ptd: std::ptr::null_mut(),
                dwAspect: 1,
                lindex: -1,
                tymed: TYMED_HGLOBAL.0 as u32,
            },
            FORMATETC {
                cfFormat: cf_paste_succeeded(),
                ptd: std::ptr::null_mut(),
                dwAspect: 1,
                lindex: -1,
                tymed: TYMED_HGLOBAL.0 as u32,
            },
        ];
        let en = NativeFormatEtcEnum::new(fmts);
        Ok(en.into())
    }
    /// 不支持数据变更通知订阅。 / Data advisory connection is not supported.
    fn DAdvise(
        &self,
        _a: *const FORMATETC,
        _b: u32,
        _c: windows_core::Ref<'_, windows::Win32::System::Com::IAdviseSink>,
    ) -> WinResult<u32> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
    /// 不支持取消通知订阅。 / Unadvise is not supported.
    fn DUnadvise(&self, _dwconnection: u32) -> WinResult<()> { Err(windows_core::Error::from(E_NOTIMPL)) }
    /// 不支持枚举通知连接。 / EnumDAdvise is not supported.
    fn EnumDAdvise(&self) -> WinResult<windows::Win32::System::Com::IEnumSTATDATA> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
}

impl IDropSource_Impl for NativeDragObject_Impl {
    /// 根据鼠标与 ESC 状态判断继续/放下/取消。 / Decide continue/drop/cancel by key state.
    fn QueryContinueDrag(&self, fescapepressed: BOOL, grfkeystate: MODIFIERKEYS_FLAGS) -> HRESULT {
        if fescapepressed.as_bool() {
            return DRAGDROP_S_CANCEL;
        }
        if (grfkeystate.0 & 0x0001) == 0 {
            return DRAGDROP_S_DROP;
        }
        S_OK
    }
    /// 使用系统默认拖拽光标反馈。 / Use system default drag feedback cursors.
    fn GiveFeedback(&self, _dweffect: DROPEFFECT) -> HRESULT {
        DRAGDROP_S_USEDEFAULTCURSORS
    }
}

#[implement(IDataObject, IDropSource)]
struct NativeVirtualFileDragObject {
    source_path: String,
    display_name: String,
    sftp_writer: Option<Arc<Mutex<Box<dyn Write + Send>>>>,
    sftp_command_lock: Option<Arc<Mutex<()>>>,
    remote_name: Option<String>,
}

struct LazyStreamState {
    started: bool,
    ready: bool,
    failed: bool,
    file_len: u64,
}

#[implement(IStream)]
struct NativeLazyContentStream {
    source_path: String,
    sftp_writer: Option<Arc<Mutex<Box<dyn Write + Send>>>>,
    sftp_command_lock: Option<Arc<Mutex<()>>>,
    remote_name: Option<String>,
    state: Arc<(Mutex<LazyStreamState>, Condvar)>,
    position: Mutex<usize>,
}

impl NativeVirtualFileDragObject {
    /// 获取 FileGroupDescriptorW 的格式 ID。 / Get format id for FileGroupDescriptorW.
    fn cf_file_descriptor() -> u16 {
        unsafe { RegisterClipboardFormatW(w!("FileGroupDescriptorW")) as u16 }
    }

    /// 获取 FileContents 的格式 ID。 / Get format id for FileContents.
    fn cf_file_contents() -> u16 {
        unsafe { RegisterClipboardFormatW(w!("FileContents")) as u16 }
    }
}

impl NativeLazyContentStream {
    fn new(
        source_path: String,
        sftp_writer: Option<Arc<Mutex<Box<dyn Write + Send>>>>,
        sftp_command_lock: Option<Arc<Mutex<()>>>,
        remote_name: Option<String>,
    ) -> Self {
        Self {
            source_path,
            sftp_writer,
            sftp_command_lock,
            remote_name,
            state: Arc::new((
                Mutex::new(LazyStreamState {
                    started: false,
                    ready: false,
                    failed: false,
                    file_len: 0,
                }),
                Condvar::new(),
            )),
            position: Mutex::new(0),
        }
    }

    fn ensure_started(&self) {
        let (lock, _cv) = &*self.state;
        let mut st = match lock.lock() {
            Ok(v) => v,
            Err(_) => return,
        };
        if st.started {
            return;
        }
        st.started = true;
        drop(st);

        let state = Arc::clone(&self.state);
        let source_path = self.source_path.clone();
        let writer = self.sftp_writer.clone();
        let command_lock = self.sftp_command_lock.clone();
        let remote_name = self.remote_name.clone();
        thread::spawn(move || {
            let result = ensure_virtual_drag_source_file(
                &source_path,
                writer.as_ref(),
                command_lock.as_ref(),
                remote_name.as_deref(),
            )
            .and_then(|_| {
                #[cfg(target_os = "windows")]
                {
                    let cpath = std::ffi::CString::new(source_path.clone())
                        .map_err(|e| format!("构造路径失败: {e}"))?;
                    let size = unsafe { mmshell_get_file_size_utf8(cpath.as_ptr()) };
                    if size == 0 {
                        return Err(format!("C++ 获取文件大小失败: {}", source_path));
                    }
                    Ok(size)
                }
                #[cfg(not(target_os = "windows"))]
                {
                    std::fs::metadata(&source_path)
                        .map(|m| m.len())
                        .map_err(|e| format!("读取下载文件元数据失败 {}: {e}", source_path))
                }
            });
            let (lock, cv) = &*state;
            if let Ok(mut st) = lock.lock() {
                match result {
                    Ok(file_len) => {
                        st.file_len = file_len;
                        st.ready = true;
                        st.failed = false;
                    }
                    Err(err) => {
                        eprintln!("[mmshell][virtual-lazy-stream] prepare failed source={} err={}", source_path, err);
                        st.ready = false;
                        st.failed = true;
                        st.file_len = 0;
                    }
                }
                cv.notify_all();
            }
        });
    }
}

impl ISequentialStream_Impl for NativeLazyContentStream_Impl {
    fn Read(&self, pv: *mut c_void, cb: u32, pcbread: *mut u32) -> HRESULT {
        if pv.is_null() {
            return E_UNEXPECTED;
        }
        self.ensure_started();
        let (lock, cv) = &*self.state;
        let mut st = match lock.lock() {
            Ok(v) => v,
            Err(_) => return E_UNEXPECTED,
        };
        while !st.ready && !st.failed {
            st = match cv.wait(st) {
                Ok(v) => v,
                Err(_) => return E_UNEXPECTED,
            };
        }
        if st.failed {
            if !pcbread.is_null() {
                unsafe { *pcbread = 0; }
            }
            // 返回空读取（而非抛系统错误），避免 Explorer 弹灾难性故障。
            return S_OK;
        }
        let mut read_size: usize = 0;
        if st.ready {
            let mut pos = match self.position.lock() {
                Ok(v) => v,
                Err(_) => return E_UNEXPECTED,
            };
            if *pos < st.file_len as usize {
                let remain = (st.file_len as usize) - *pos;
                read_size = remain.min(cb as usize);
                #[cfg(target_os = "windows")]
                {
                    let cpath = match std::ffi::CString::new(self.source_path.clone()) {
                        Ok(v) => v,
                        Err(_) => {
                            if !pcbread.is_null() { unsafe { *pcbread = 0; } }
                            return S_OK;
                        }
                    };
                    let mut out_read: u32 = 0;
                    let rc = unsafe {
                        mmshell_read_file_chunk_utf8(
                            cpath.as_ptr(),
                            *pos as u64,
                            read_size as u32,
                            pv as *mut u8,
                            &mut out_read as *mut u32,
                        )
                    };
                    if rc == 0 {
                        read_size = out_read as usize;
                        *pos += read_size;
                    } else {
                        read_size = 0;
                    }
                }
                #[cfg(not(target_os = "windows"))]
                {
                    let mut file = match std::fs::File::open(&self.source_path) {
                        Ok(f) => f,
                        Err(_) => {
                            if !pcbread.is_null() {
                                unsafe { *pcbread = 0; }
                            }
                            return S_OK;
                        }
                    };
                    if file.seek(SeekFrom::Start(*pos as u64)).is_err() {
                        if !pcbread.is_null() {
                            unsafe { *pcbread = 0; }
                        }
                        return S_OK;
                    }
                    let mut buf = vec![0u8; read_size];
                    let n = match file.read(&mut buf) {
                        Ok(n) => n,
                        Err(_) => 0,
                    };
                    read_size = n;
                    if read_size > 0 {
                        unsafe {
                            std::ptr::copy_nonoverlapping(buf.as_ptr(), pv as *mut u8, read_size);
                        }
                        *pos += read_size;
                    }
                }
            }
        }
        if !pcbread.is_null() {
            unsafe { *pcbread = read_size as u32; }
        }
        S_OK
    }

    fn Write(&self, _pv: *const c_void, _cb: u32, _pcbwritten: *mut u32) -> HRESULT {
        E_NOTIMPL
    }
}

impl IStream_Impl for NativeLazyContentStream_Impl {
    fn Seek(&self, dlibmove: i64, dworigin: STREAM_SEEK, plibnewposition: *mut u64) -> WinResult<()> {
        let (lock, _) = &*self.state;
        let st = lock
            .lock()
            .map_err(|_| windows_core::Error::new(E_UNEXPECTED, "lazy stream state poisoned"))?;
        let mut pos = self
            .position
            .lock()
            .map_err(|_| windows_core::Error::new(E_UNEXPECTED, "lazy stream position poisoned"))?;
        let base: i64 = match dworigin {
            STREAM_SEEK(0) => 0,
            STREAM_SEEK(1) => *pos as i64,
            STREAM_SEEK(2) => st.file_len as i64,
            _ => return Err(windows_core::Error::new(E_FAIL, "invalid seek origin")),
        };
        let new_pos = base.saturating_add(dlibmove).max(0) as usize;
        *pos = new_pos;
        if !plibnewposition.is_null() {
            unsafe { *plibnewposition = *pos as u64; }
        }
        Ok(())
    }

    fn SetSize(&self, _libnewsize: u64) -> WinResult<()> { Err(windows_core::Error::from(E_NOTIMPL)) }
    fn CopyTo(&self, _pstm: windows_core::Ref<'_, IStream>, _cb: u64, _pcbread: *mut u64, _pcbwritten: *mut u64) -> WinResult<()> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
    fn Commit(&self, _grfcommitflags: &STGC) -> WinResult<()> { Ok(()) }
    fn Revert(&self) -> WinResult<()> { Err(windows_core::Error::from(E_NOTIMPL)) }
    fn LockRegion(&self, _liboffset: u64, _cb: u64, _dwlocktype: &LOCKTYPE) -> WinResult<()> { Err(windows_core::Error::from(E_NOTIMPL)) }
    fn UnlockRegion(&self, _liboffset: u64, _cb: u64, _dwlocktype: u32) -> WinResult<()> { Err(windows_core::Error::from(E_NOTIMPL)) }
    fn Stat(&self, pstatstg: *mut STATSTG, _grfstatflag: &STATFLAG) -> WinResult<()> {
        if pstatstg.is_null() {
            return Err(windows_core::Error::new(E_UNEXPECTED, "STATSTG is null"));
        }
        let (lock, _) = &*self.state;
        let st = lock
            .lock()
            .map_err(|_| windows_core::Error::new(E_UNEXPECTED, "lazy stream state poisoned"))?;
        unsafe {
            let mut stat = STATSTG::default();
            stat.r#type = STGTY_STREAM.0 as u32;
            stat.cbSize = st.file_len;
            *pstatstg = stat;
        }
        Ok(())
    }
    fn Clone(&self) -> WinResult<IStream> {
        let cloned = NativeLazyContentStream::new(
            self.source_path.clone(),
            self.sftp_writer.clone(),
            self.sftp_command_lock.clone(),
            self.remote_name.clone(),
        );
        Ok(cloned.into())
    }
}

/// 构造 FILEGROUPDESCRIPTORW 数据块。 / Build FILEGROUPDESCRIPTORW HGLOBAL payload.
fn build_file_group_descriptor_hglobal(display_name: &str, file_len: u64) -> Result<HGLOBAL, String> {
    let size = std::mem::size_of::<FILEGROUPDESCRIPTORW>();
    let hglobal = unsafe { GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, size) }
        .map_err(|e| format!("GlobalAlloc(FGD) 失败: {e}"))?;
    let locked = unsafe { GlobalLock(hglobal) } as *mut u8;
    if locked.is_null() {
        return Err("GlobalLock(FGD) 失败".to_string());
    }
    unsafe {
        let group = locked as *mut FILEGROUPDESCRIPTORW;
        (*group).cItems = 1;
        let item = (*group).fgd.as_mut_ptr();
        let mut desc = FILEDESCRIPTORW::default();
        desc.dwFlags = (FD_FILESIZE.0 | FD_ATTRIBUTES.0 | FD_PROGRESSUI.0) as u32;
        desc.dwFileAttributes = std::fs::metadata(&PathBuf::from(display_name))
            .ok()
            .map(|m| if m.is_dir() { 0x10 } else { 0x80 })
            .unwrap_or(0x80);
        desc.nFileSizeHigh = (file_len >> 32) as u32;
        desc.nFileSizeLow = (file_len & 0xFFFF_FFFF) as u32;
        const FILE_NAME_CAP: usize = 260;
        let mut wide_name: Vec<u16> = display_name.encode_utf16().collect();
        if wide_name.len() >= FILE_NAME_CAP {
            wide_name.truncate(FILE_NAME_CAP - 1);
        }
        let mut file_name_buf = [0u16; FILE_NAME_CAP];
        for (i, ch) in wide_name.iter().enumerate() {
            file_name_buf[i] = *ch;
        }
        desc.cFileName = file_name_buf;
        *item = desc;
        let _ = GlobalUnlock(hglobal);
    }
    Ok(hglobal)
}

/// 读取本地文件并构造 FileContents 数据块。 / Read local file and build FileContents HGLOBAL.
fn build_file_contents_hglobal(source_path: &str) -> Result<HGLOBAL, String> {
    let data = std::fs::read(source_path).map_err(|e| format!("读取拖拽源文件失败: {e}"))?;
    let hglobal = unsafe { GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, data.len()) }
        .map_err(|e| format!("GlobalAlloc(FileContents) 失败: {e}"))?;
    let locked = unsafe { GlobalLock(hglobal) } as *mut u8;
    if locked.is_null() {
        return Err("GlobalLock(FileContents) 失败".to_string());
    }
    unsafe {
        std::ptr::copy_nonoverlapping(data.as_ptr(), locked, data.len());
        let _ = GlobalUnlock(hglobal);
    }
    Ok(hglobal)
}

/// 构造最小空内容 HGLOBAL，避免在拖拽取数阶段抛出灾难性 COM 错误。 / Build minimal empty HGLOBAL to avoid catastrophic COM failure in FileContents stage.
fn build_empty_contents_hglobal() -> Result<HGLOBAL, String> {
    let hglobal = unsafe { GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, 1) }
        .map_err(|e| format!("GlobalAlloc(EmptyFileContents) 失败: {e}"))?;
    Ok(hglobal)
}

/// 基于 HGLOBAL 构造 IStream，用于 FileContents(ISTREAM) 传输。 / Build IStream from HGLOBAL for FileContents(ISTREAM).
fn build_file_contents_istream(source_path: &str) -> Result<windows::Win32::System::Com::IStream, String> {
    let hglobal = build_file_contents_hglobal(source_path)?;
    let stream = unsafe { CreateStreamOnHGlobal(hglobal, true) }
        .map_err(|e| format!("CreateStreamOnHGlobal 失败: {e}"))?;
    Ok(stream)
}

/// 对 SFTP 参数做安全引号转义。 / Quote and escape SFTP argument safely.
fn quote_sftp_arg(raw: &str) -> String {
    let escaped = raw.replace('\\', "\\\\").replace('"', "\\\"");
    format!("\"{escaped}\"")
}

/// 等待文件或目录在超时时间内出现。 / Wait until file/dir exists within timeout.
fn wait_for_file_ready(path: &Path, timeout_ms: u64) -> bool {
    let start = std::time::Instant::now();
    while start.elapsed().as_millis() < timeout_ms as u128 {
        if path.exists() {
            return true;
        }
        thread::sleep(Duration::from_millis(120));
    }
    false
}

fn spawn_sftp_background_download(
    writer: Arc<Mutex<Box<dyn Write + Send>>>,
    command_lock: Arc<Mutex<()>>,
    remote_path: String,
    local_path: PathBuf,
    done_marker: PathBuf,
    err_marker: PathBuf,
) {
    thread::spawn(move || {
        let _ = std::fs::remove_file(&done_marker);
        let _ = std::fs::remove_file(&err_marker);
        let result: Result<(), String> = (|| {
            if let Some(parent) = local_path.parent() {
                std::fs::create_dir_all(parent)
                    .map_err(|e| format!("创建缓存目录失败 {}: {e}", parent.display()))?;
            }
            let cmd = format!(
                "get {} {}\r",
                quote_sftp_arg(&remote_path),
                quote_sftp_arg(&local_path.to_string_lossy())
            );
            let _cmd_guard = command_lock
                .lock()
                .map_err(|_| "sftp command lock poisoned".to_string())?;
            {
                let mut lock = writer
                    .lock()
                    .map_err(|_| "SFTP writer lock poisoned".to_string())?;
                lock.write_all(cmd.as_bytes())
                    .map_err(|e| format!("发送 SFTP 下载命令失败: {e}"))?;
                lock.flush().map_err(|e| format!("刷新 SFTP writer 失败: {e}"))?;
            }
            // 只要文件出现即允许流式读取；完成标记用于通知 EOF。
            if !wait_for_file_ready(&local_path, 300_000) {
                return Err(format!("流式下载超时: {}", local_path.display()));
            }
            // 等待文件稳定后落 done 标记，通知流结束。
            let _ = wait_for_file_stable(&local_path, 600_000, 1200);
            Ok(())
        })();

        match result {
            Ok(()) => {
                let _ = std::fs::write(&done_marker, b"ok");
            }
            Err(err) => {
                let _ = std::fs::write(&err_marker, err.as_bytes());
            }
        }
    });
}

/// 等待文件大小稳定，避免刚创建就被误判为下载完成。 / Wait until file size is stable to avoid premature completion.
fn wait_for_file_stable(path: &Path, timeout_ms: u64, stable_ms: u64) -> bool {
    let start = std::time::Instant::now();
    let mut last_size: Option<u64> = None;
    let mut stable_since: Option<std::time::Instant> = None;
    while start.elapsed().as_millis() < timeout_ms as u128 {
        if let Ok(meta) = std::fs::metadata(path) {
            let size = meta.len();
            if Some(size) == last_size {
                if let Some(since) = stable_since {
                    if since.elapsed().as_millis() >= stable_ms as u128 {
                        return true;
                    }
                } else {
                    stable_since = Some(std::time::Instant::now());
                }
            } else {
                last_size = Some(size);
                stable_since = Some(std::time::Instant::now());
            }
        }
        thread::sleep(Duration::from_millis(120));
    }
    false
}

/// 按需确保虚拟拖拽源文件已在本地可读。 / Ensure local source exists for virtual drag on demand.
fn ensure_virtual_drag_source_file(
    source_path: &str,
    sftp_writer: Option<&Arc<Mutex<Box<dyn Write + Send>>>>,
    sftp_command_lock: Option<&Arc<Mutex<()>>>,
    remote_name: Option<&str>,
) -> Result<(), String> {
    let local = PathBuf::from(source_path);
    if local.exists() {
        return Ok(());
    }
    let writer = sftp_writer.ok_or_else(|| "缺少 SFTP 会话，无法按需下载".to_string())?;
    let remote = remote_name.ok_or_else(|| "缺少远端文件名，无法按需下载".to_string())?;
    if let Some(parent) = local.parent() {
        std::fs::create_dir_all(parent)
            .map_err(|e| format!("创建本地缓存目录失败 {}: {e}", parent.display()))?;
    }
    let cmd = format!("get {} {}\r", quote_sftp_arg(remote), quote_sftp_arg(source_path));
    let _cmd_guard = sftp_command_lock
        .ok_or_else(|| "缺少 SFTP 命令锁，无法按需下载".to_string())?
        .lock()
        .map_err(|_| "sftp command lock poisoned".to_string())?;
    {
        let mut lock = writer
            .lock()
            .map_err(|_| "SFTP writer lock poisoned".to_string())?;
        lock.write_all(cmd.as_bytes())
            .map_err(|e| format!("发送 SFTP 下载命令失败: {e}"))?;
        lock.flush().map_err(|e| format!("刷新 SFTP writer 失败: {e}"))?;
    }
    if !wait_for_file_ready(&local, 30000) {
        return Err(format!("按需下载超时: {}", local.display()));
    }
    Ok(())
}

/// 将远端 SFTP 文件/目录下载到系统临时目录。 / Download remote SFTP file/dir into temp directory.
#[allow(dead_code)]
fn download_sftp_to_temp_path(
    writer: &Arc<Mutex<Box<dyn Write + Send>>>,
    command_lock: &Arc<Mutex<()>>,
    remote_path: &str,
    display_name: &str,
    is_dir: bool,
) -> Result<PathBuf, String> {
    let base_name = sanitize_placeholder_name(display_name);
    let cache_root = resolve_drag_cache_dir()?;
    std::fs::create_dir_all(&cache_root)
        .map_err(|e| format!("创建临时目录失败 {}: {e}", cache_root.display()))?;
    // 每次传输使用唯一目录，避免同名文件被占用时触发 Windows 重试/取消弹窗。
    let ts = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_millis())
        .unwrap_or_default();
    let op_dir = cache_root.join(format!("job-{ts}"));
    std::fs::create_dir_all(&op_dir)
        .map_err(|e| format!("创建传输目录失败 {}: {e}", op_dir.display()))?;
    let local_path = op_dir.join(&base_name);
    let cmd = if is_dir {
        format!(
            "get -r {} {}\r",
            quote_sftp_arg(remote_path),
            quote_sftp_arg(&op_dir.to_string_lossy())
        )
    } else {
        format!(
            "get {} {}\r",
            quote_sftp_arg(remote_path),
            quote_sftp_arg(&local_path.to_string_lossy())
        )
    };
    eprintln!(
        "[mmshell][native-drag-hdrop] download begin remote={} local={} is_dir={} cmd={}",
        remote_path,
        local_path.display(),
        is_dir,
        cmd.trim()
    );
    let _cmd_guard = command_lock
        .lock()
        .map_err(|_| "sftp command lock poisoned".to_string())?;
    let expected = if is_dir {
        if local_path.exists() { local_path.clone() } else { op_dir.clone() }
    } else {
        local_path.clone()
    };
    {
        let mut lock = writer
            .lock()
            .map_err(|_| "SFTP writer lock poisoned".to_string())?;
        lock.write_all(cmd.as_bytes())
            .map_err(|e| format!("发送 SFTP 下载命令失败: {e}"))?;
        lock.flush().map_err(|e| format!("刷新 SFTP writer 失败: {e}"))?;
    }
    let timeout_ms = if is_dir { 240000 } else { 120000 };
    if !wait_for_file_ready(&expected, timeout_ms) {
        eprintln!(
            "[mmshell][native-drag-hdrop] download timeout remote={} expected={} exists_now={}",
            remote_path,
            expected.display(),
            expected.exists()
        );
        return Err(format!("下载到临时目录超时: {}", expected.display()));
    }
    if expected.is_file() && !wait_for_file_stable(&expected, 30_000, 1_200) {
        // 稳定性检测仅做保护，不阻断拖拽主流程；否则会出现“进度有了但拖拽没启动”。
        eprintln!(
            "[mmshell][native-drag-hdrop] file stability check not reached, continue anyway file={}",
            expected.display()
        );
    }
    if expected.is_file() {
        match std::fs::metadata(&expected) {
            Ok(meta) => {
                eprintln!(
                    "[mmshell][native-drag-hdrop] download ready file={} size={} bytes",
                    expected.display(),
                    meta.len()
                );
            }
            Err(err) => {
                eprintln!(
                    "[mmshell][native-drag-hdrop] download ready but metadata failed file={} err={}",
                    expected.display(),
                    err
                );
            }
        }
    } else {
        eprintln!(
            "[mmshell][native-drag-hdrop] download ready dir={}",
            expected.display()
        );
    }
    Ok(expected)
}

/// 使用独立一次性 sftp 进程下载，避免复用会话导致 busy 冲突。 / Download via one-shot dedicated sftp process to avoid busy conflicts.
#[allow(dead_code)]
fn download_sftp_to_temp_path_one_shot(
    address: &str,
    password: &str,
    remote_path: &str,
    display_name: &str,
    is_dir: bool,
) -> Result<PathBuf, String> {
    let parsed = parse_address(address)?;
    let base_name = sanitize_placeholder_name(display_name);
    let temp_dir = resolve_drag_cache_dir()?;
    std::fs::create_dir_all(&temp_dir)
        .map_err(|e| format!("创建临时目录失败 {}: {e}", temp_dir.display()))?;
    let local_path = temp_dir.join(&base_name);
    if local_path.exists() {
        if local_path.is_dir() {
            let _ = std::fs::remove_dir_all(&local_path);
        } else {
            let _ = std::fs::remove_file(&local_path);
        }
    }
    let get_cmd = if is_dir {
        format!(
            "get -r {} {}\r",
            quote_sftp_arg(remote_path),
            quote_sftp_arg(&temp_dir.to_string_lossy())
        )
    } else {
        format!(
            "get {} {}\r",
            quote_sftp_arg(remote_path),
            quote_sftp_arg(&local_path.to_string_lossy())
        )
    };

    eprintln!(
        "[mmshell][native-drag-hdrop] one-shot download begin remote={} local={} is_dir={}",
        remote_path,
        local_path.display(),
        is_dir
    );
    let pty_system = native_pty_system();
    let pty_pair = pty_system
        .openpty(PtySize {
            rows: 24,
            cols: 120,
            pixel_width: 0,
            pixel_height: 0,
        })
        .map_err(|err| format!("创建 one-shot SFTP PTY 失败: {err}"))?;

    let mut cmd = CommandBuilder::new("sftp.exe");
    cmd.arg("-o");
    cmd.arg("StrictHostKeyChecking=accept-new");
    cmd.arg("-o");
    cmd.arg("HostKeyAlgorithms=+ssh-rsa");
    cmd.arg("-o");
    cmd.arg("PubkeyAcceptedAlgorithms=+ssh-rsa");
    cmd.arg("-P");
    cmd.arg(parsed.port.to_string());
    cmd.arg(format!("{}@{}", parsed.user, parsed.host));

    let mut child = pty_pair
        .slave
        .spawn_command(cmd)
        .map_err(|err| format!("启动 one-shot sftp.exe 失败: {err}"))?;
    drop(pty_pair.slave);
    let reader = pty_pair
        .master
        .try_clone_reader()
        .map_err(|err| format!("创建 one-shot SFTP reader 失败: {err}"))?;
    let mut writer = pty_pair
        .master
        .take_writer()
        .map_err(|err| format!("创建 one-shot SFTP writer 失败: {err}"))?;

    let expected = if is_dir {
        if local_path.exists() { local_path.clone() } else { temp_dir.clone() }
    } else {
        local_path.clone()
    };

    let mut sent_password = false;
    let mut sent_get = false;
    let mut sent_bye = false;
    let mut transfer_done_prompt = false;
    let mut reader = BufReader::new(reader);
    let mut buf = [0_u8; 4096];
    let started = std::time::Instant::now();
    let hard_timeout_ms = if is_dir { 120000 } else { 60000 };

    loop {
        if started.elapsed().as_millis() > hard_timeout_ms as u128 {
            let _ = child.kill();
            return Err("one-shot sftp 下载超时".to_string());
        }
        let n = reader
            .read(&mut buf)
            .map_err(|e| format!("读取 one-shot sftp 输出失败: {e}"))?;
        if n == 0 {
            continue;
        }
        let text = String::from_utf8_lossy(&buf[..n]).to_string();
        let lower = text.to_lowercase();
        if lower.contains("permission denied") || lower.contains("couldn't") || lower.contains("failure") {
            let _ = child.kill();
            return Err(format!("one-shot sftp 失败: {}", text.trim()));
        }
        if !sent_password && lower.contains("password:") {
            writer
                .write_all(format!("{password}\r").as_bytes())
                .map_err(|e| format!("发送 one-shot 密码失败: {e}"))?;
            writer.flush().map_err(|e| format!("刷新 one-shot 密码失败: {e}"))?;
            sent_password = true;
            continue;
        }
        if lower.contains("sftp>") && !sent_get {
            writer
                .write_all(get_cmd.as_bytes())
                .map_err(|e| format!("发送 one-shot get 命令失败: {e}"))?;
            writer.flush().map_err(|e| format!("刷新 one-shot get 失败: {e}"))?;
            sent_get = true;
            continue;
        }
        if sent_get && !sent_bye && lower.contains("sftp>") {
            transfer_done_prompt = true;
            writer
                .write_all(b"bye\r")
                .map_err(|e| format!("发送 one-shot bye 失败: {e}"))?;
            writer.flush().map_err(|e| format!("刷新 one-shot bye 失败: {e}"))?;
            sent_bye = true;
        }
        if transfer_done_prompt {
            if expected.is_file() {
                if wait_for_file_stable(&expected, 20000, 1200) {
                    eprintln!(
                        "[mmshell][native-drag-hdrop] one-shot file stable expected={}",
                        expected.display()
                    );
                    break;
                }
            } else if wait_for_file_ready(&expected, 20000) {
                eprintln!(
                    "[mmshell][native-drag-hdrop] one-shot dir ready expected={}",
                    expected.display()
                );
                break;
            }
        }
    }
    let _ = child.kill();
    Ok(expected)
}

/// 自动识别本地拖拽缓存目录。 / Auto-detect local cache directory for native drag.
fn resolve_drag_cache_dir() -> Result<PathBuf, String> {
    if let Ok(exe) = std::env::current_exe() {
        if let Some(exe_dir) = exe.parent() {
            let candidate = exe_dir.join("drag-cache");
            if std::fs::create_dir_all(&candidate).is_ok() {
                eprintln!(
                    "[mmshell][native-drag-hdrop] cache dir auto-detected from exe: {}",
                    candidate.display()
                );
                return Ok(candidate);
            }
        }
    }
    if let Ok(local_app_data) = std::env::var("LOCALAPPDATA") {
        let fallback = PathBuf::from(local_app_data).join("MMShell").join("drag-cache");
        std::fs::create_dir_all(&fallback)
            .map_err(|e| format!("创建缓存目录失败 {}: {e}", fallback.display()))?;
        eprintln!(
            "[mmshell][native-drag-hdrop] cache dir fallback LOCALAPPDATA: {}",
            fallback.display()
        );
        return Ok(fallback);
    }
    let temp = std::env::temp_dir().join("mmshell_sftp_drag");
    std::fs::create_dir_all(&temp)
        .map_err(|e| format!("创建缓存目录失败 {}: {e}", temp.display()))?;
    eprintln!(
        "[mmshell][native-drag-hdrop] cache dir fallback TEMP: {}",
        temp.display()
    );
    Ok(temp)
}

/// 转换为以 null 结尾的 UTF-16 字符串。 / Convert string to null-terminated UTF-16.
fn to_wide_null(s: &str) -> Vec<u16> {
    let mut v: Vec<u16> = s.encode_utf16().collect();
    v.push(0);
    v
}

/// 使用 Shell API 由路径创建系统 IDataObject。 / Create system IDataObject from path via Shell APIs.
fn create_shell_data_object_from_path(path: &Path) -> Result<IDataObject, String> {
    let path_canonical = path
        .canonicalize()
        .map_err(|e| format!("路径规范化失败 {}: {e}", path.display()))?;
    let path_str = path_canonical.to_string_lossy().to_string();
    let wide = to_wide_null(&path_str);
    let mut abs_pidl: *mut Common::ITEMIDLIST = std::ptr::null_mut();
    let parsed = unsafe {
        SHParseDisplayName(
            windows_core::PCWSTR(wide.as_ptr()),
            None,
            &mut abs_pidl,
            0,
            None,
        )
    };
    parsed.map_err(|e| format!("SHParseDisplayName 失败: {e}"))?;
    if abs_pidl.is_null() {
        return Err("SHParseDisplayName 返回空 PIDL".to_string());
    }

    let child_array: [*const Common::ITEMIDLIST; 1] = [abs_pidl as *const Common::ITEMIDLIST];
    let obj = unsafe {
        SHCreateDataObject(
            None,
            Some(&child_array),
            None::<&IDataObject>,
        )
    }
    .map_err(|e| format!("SHCreateDataObject 失败: {e}"))?;
    let parse_cleanup = unsafe {
        windows::Win32::System::Com::CoTaskMemFree(Some(abs_pidl as *const core::ffi::c_void));
    };
    let _ = parse_cleanup;
    Ok(obj)
}

impl IDataObject_Impl for NativeVirtualFileDragObject_Impl {
    /// 提供虚拟文件描述或内容数据。 / Provide virtual file descriptor/content payloads.
    fn GetData(&self, pformatetcin: *const FORMATETC) -> WinResult<STGMEDIUM> {
        if pformatetcin.is_null() {
            return Err(windows_core::Error::new(E_UNEXPECTED, "FORMATETC is null"));
        }
        let fmt = unsafe { *pformatetcin };
        let cf_descriptor = NativeVirtualFileDragObject::cf_file_descriptor();
        let cf_contents = NativeVirtualFileDragObject::cf_file_contents();

        let mut medium = STGMEDIUM::default();
        if fmt.cfFormat == cf_descriptor && (fmt.tymed & TYMED_HGLOBAL.0 as u32) != 0 {
            medium.tymed = TYMED_HGLOBAL.0 as u32;
            let file_len = std::fs::metadata(&self.source_path).map(|m| m.len()).unwrap_or(0);
            medium.u.hGlobal = build_file_group_descriptor_hglobal(&self.display_name, file_len)
                .map_err(|e| windows_core::Error::new(E_UNEXPECTED, e))?;
            medium.pUnkForRelease = ManuallyDrop::new(None);
            return Ok(medium);
        }
        if fmt.cfFormat == cf_contents
            && ((fmt.tymed & TYMED_ISTREAM.0 as u32) != 0 || (fmt.tymed & TYMED_HGLOBAL.0 as u32) != 0)
        {
            let ensure_result = ensure_virtual_drag_source_file(
                &self.source_path,
                self.sftp_writer.as_ref(),
                self.sftp_command_lock.as_ref(),
                self.remote_name.as_deref(),
            );
            if let Err(err) = &ensure_result {
                eprintln!(
                    "[mmshell][virtual] FileContents prepare failed, fallback empty payload source={} err={}",
                    self.source_path,
                    err
                );
            }
            if (fmt.tymed & TYMED_ISTREAM.0 as u32) != 0 {
                medium.tymed = TYMED_ISTREAM.0 as u32;
                let stream = if ensure_result.is_ok() {
                    build_file_contents_istream(&self.source_path)
                        .map_err(|e| windows_core::Error::new(E_UNEXPECTED, e))?
                } else {
                    let hglobal = build_empty_contents_hglobal()
                        .map_err(|e| windows_core::Error::new(E_UNEXPECTED, e))?;
                    unsafe { CreateStreamOnHGlobal(hglobal, true) }
                        .map_err(|e| windows_core::Error::new(E_UNEXPECTED, format!("CreateStreamOnHGlobal(Empty) 失败: {e}")))?
                };
                medium.u.pstm = ManuallyDrop::new(Some(stream));
            } else {
                medium.tymed = TYMED_HGLOBAL.0 as u32;
                medium.u.hGlobal = if ensure_result.is_ok() {
                    build_file_contents_hglobal(&self.source_path)
                        .map_err(|e| windows_core::Error::new(E_UNEXPECTED, e))?
                } else {
                    build_empty_contents_hglobal()
                        .map_err(|e| windows_core::Error::new(E_UNEXPECTED, e))?
                };
            }
            medium.pUnkForRelease = ManuallyDrop::new(None);
            return Ok(medium);
        }
        if fmt.cfFormat == cf_preferred_drop_effect() && (fmt.tymed & TYMED_HGLOBAL.0 as u32) != 0 {
            medium.tymed = TYMED_HGLOBAL.0 as u32;
            medium.u.hGlobal = build_drop_effect_hglobal(DROPEFFECT_COPY.0 as u32)
                .map_err(|e| windows_core::Error::new(E_UNEXPECTED, e))?;
            medium.pUnkForRelease = ManuallyDrop::new(None);
            return Ok(medium);
        }
        Err(windows_core::Error::new(DV_E_FORMATETC, "unsupported format"))
    }

    /// 不支持 GetDataHere。 / GetDataHere is not supported for virtual drag object.
    fn GetDataHere(&self, _pformatetc: *const FORMATETC, _pmedium: *mut STGMEDIUM) -> WinResult<()> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
    /// 检查虚拟拖拽支持的格式。 / Check supported formats for virtual drag.
    fn QueryGetData(&self, pformatetc: *const FORMATETC) -> HRESULT {
        if pformatetc.is_null() {
            return DV_E_FORMATETC;
        }
        let fmt = unsafe { *pformatetc };
        let cf_descriptor = NativeVirtualFileDragObject::cf_file_descriptor();
        let cf_contents = NativeVirtualFileDragObject::cf_file_contents();
        if fmt.cfFormat == cf_descriptor && (fmt.tymed & TYMED_HGLOBAL.0 as u32) != 0 {
            return S_OK;
        }
        if fmt.cfFormat == cf_contents
            && ((fmt.tymed & TYMED_ISTREAM.0 as u32) != 0 || (fmt.tymed & TYMED_HGLOBAL.0 as u32) != 0)
        {
            return S_OK;
        }
        if fmt.cfFormat == cf_preferred_drop_effect() && (fmt.tymed & TYMED_HGLOBAL.0 as u32) != 0 {
            return S_OK;
        }
        DV_E_FORMATETC
    }
    /// 不提供规范化格式映射。 / Canonical format mapping is not implemented.
    fn GetCanonicalFormatEtc(&self, _a: *const FORMATETC, _b: *mut FORMATETC) -> HRESULT { E_NOTIMPL }
    /// 接受目标端回写的拖放结果格式，提升 Explorer 兼容性。 / Accept target-side drop-result formats for Explorer compatibility.
    fn SetData(&self, a: *const FORMATETC, _b: *const STGMEDIUM, _c: BOOL) -> WinResult<()> {
        if a.is_null() {
            return Err(windows_core::Error::new(E_UNEXPECTED, "FORMATETC is null"));
        }
        let fmt = unsafe { *a };
        if fmt.cfFormat == cf_performed_drop_effect() {
            eprintln!("[mmshell][virtual] SetData: PerformedDropEffect received");
            return Ok(());
        }
        if fmt.cfFormat == cf_paste_succeeded() {
            eprintln!("[mmshell][virtual] SetData: PasteSucceeded received");
            return Ok(());
        }
        if fmt.cfFormat == cf_preferred_drop_effect() {
            eprintln!("[mmshell][virtual] SetData: PreferredDropEffect received");
            return Ok(());
        }
        eprintln!(
            "[mmshell][virtual] SetData: ignored format cf={} tymed={}",
            fmt.cfFormat,
            fmt.tymed
        );
        Ok(())
    }
    /// 枚举虚拟拖拽可提供的格式（含 FileContents 的 HGLOBAL/ISTREAM）。 / Enumerate virtual formats including FileContents HGLOBAL/ISTREAM.
    fn EnumFormatEtc(&self, dwdirection: u32) -> WinResult<windows::Win32::System::Com::IEnumFORMATETC> {
        if dwdirection != DATADIR_GET.0 as u32 {
            return Err(windows_core::Error::from(DV_E_FORMATETC));
        }
        let cf_descriptor = NativeVirtualFileDragObject::cf_file_descriptor();
        let cf_contents = NativeVirtualFileDragObject::cf_file_contents();
        let fmts = vec![
            FORMATETC {
                cfFormat: cf_descriptor,
                ptd: std::ptr::null_mut(),
                dwAspect: 1,
                lindex: -1,
                tymed: TYMED_HGLOBAL.0 as u32,
            },
            FORMATETC {
                cfFormat: cf_contents,
                ptd: std::ptr::null_mut(),
                dwAspect: 1,
                lindex: 0,
                tymed: TYMED_ISTREAM.0 as u32,
            },
            FORMATETC {
                cfFormat: cf_contents,
                ptd: std::ptr::null_mut(),
                dwAspect: 1,
                lindex: 0,
                tymed: TYMED_HGLOBAL.0 as u32,
            },
            FORMATETC {
                cfFormat: cf_preferred_drop_effect(),
                ptd: std::ptr::null_mut(),
                dwAspect: 1,
                lindex: -1,
                tymed: TYMED_HGLOBAL.0 as u32,
            },
            FORMATETC {
                cfFormat: cf_performed_drop_effect(),
                ptd: std::ptr::null_mut(),
                dwAspect: 1,
                lindex: -1,
                tymed: TYMED_HGLOBAL.0 as u32,
            },
            FORMATETC {
                cfFormat: cf_paste_succeeded(),
                ptd: std::ptr::null_mut(),
                dwAspect: 1,
                lindex: -1,
                tymed: TYMED_HGLOBAL.0 as u32,
            },
        ];
        let en = NativeFormatEtcEnum::new(fmts);
        Ok(en.into())
    }
    /// 不支持通知订阅。 / Advisory connections are not supported.
    fn DAdvise(
        &self,
        _a: *const FORMATETC,
        _b: u32,
        _c: windows_core::Ref<'_, windows::Win32::System::Com::IAdviseSink>,
    ) -> WinResult<u32> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
    /// 不支持取消通知订阅。 / Unadvise is not supported.
    fn DUnadvise(&self, _dwconnection: u32) -> WinResult<()> { Err(windows_core::Error::from(E_NOTIMPL)) }
    /// 不支持枚举通知连接。 / EnumDAdvise is not supported.
    fn EnumDAdvise(&self) -> WinResult<windows::Win32::System::Com::IEnumSTATDATA> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
}

impl IDropSource_Impl for NativeVirtualFileDragObject_Impl {
    /// 根据按键状态控制拖拽生命周期。 / Control drag lifecycle by key/button state.
    fn QueryContinueDrag(&self, fescapepressed: BOOL, grfkeystate: MODIFIERKEYS_FLAGS) -> HRESULT {
        if fescapepressed.as_bool() {
            return DRAGDROP_S_CANCEL;
        }
        if (grfkeystate.0 & 0x0001) == 0 {
            return DRAGDROP_S_DROP;
        }
        S_OK
    }
    /// 使用系统默认拖拽光标。 / Use system default drag cursors.
    fn GiveFeedback(&self, _dweffect: DROPEFFECT) -> HRESULT {
        DRAGDROP_S_USEDEFAULTCURSORS
    }
}


/// 递归复制目录到目标路径。 / Recursively copy directory to destination path.
fn copy_dir_recursive(src: &Path, dst: &Path) -> Result<(), String> {
    std::fs::create_dir_all(dst).map_err(|e| format!("创建目录失败 {}: {e}", dst.display()))?;
    let entries = std::fs::read_dir(src).map_err(|e| format!("读取目录失败 {}: {e}", src.display()))?;
    for entry in entries {
        let entry = entry.map_err(|e| format!("读取目录项失败: {e}"))?;
        let path = entry.path();
        let target_path = dst.join(entry.file_name());
        if path.is_dir() {
            copy_dir_recursive(&path, &target_path)?;
        } else {
            std::fs::copy(&path, &target_path)
                .map_err(|e| format!("复制文件失败 {} -> {}: {e}", path.display(), target_path.display()))?;
        }
    }
    Ok(())
}

/// 尝试删除本地拖拽缓存（失败仅记录日志）。 / Try deleting local drag cache (log-only on failure).
fn cleanup_drag_cache_path(path: &Path) {
    let result = if path.is_dir() {
        std::fs::remove_dir_all(path)
    } else {
        std::fs::remove_file(path)
    };
    match result {
        Ok(()) => eprintln!("[mmshell][native-drag-hdrop] cleanup ok path={}", path.display()),
        Err(err) => eprintln!(
            "[mmshell][native-drag-hdrop] cleanup failed path={} err={}",
            path.display(),
            err
        ),
    }
}

/// 延迟清理拖拽缓存，避免目标程序仍在读取时触发系统“重试/取消”弹窗。 / Delay cache cleanup to avoid Windows retry/cancel popup.
fn schedule_cleanup_drag_cache_path(path: PathBuf, delay_ms: u64) {
    thread::spawn(move || {
        thread::sleep(Duration::from_millis(delay_ms));
        cleanup_drag_cache_path(&path);
    });
}

/// M1 探针：创建占位文件并返回路径信息。 / M1 probe: create placeholder and return metadata.
#[tauri::command]
fn native_drag_out_begin(payload: NativeDragOutPayload) -> Result<NativeDragOutResult, String> {
    let cache_dir = resolve_drag_cache_dir()?;

    let ts = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_millis())
        .unwrap_or_default();

    let base = payload
        .remote_path
        .rsplit('/')
        .next()
        .map(sanitize_placeholder_name)
        .unwrap_or_else(|| "drag-item".to_string());
    let suffix = if payload.is_dir { "dir" } else { "file" };
    let placeholder_name = format!("m1-{ts}-{suffix}-{base}.txt");
    let placeholder_path = cache_dir.join(placeholder_name);

    let content = format!(
        "M1 原生拖拽占位文件\nM1 native drag placeholder file\n\nremote_path={}\nis_dir={}\ncreated_at_ms={}\n",
        payload.remote_path, payload.is_dir, ts
    );
    std::fs::write(&placeholder_path, content)
        .map_err(|err| format!("写入占位文件失败: {err}"))?;

    open_in_explorer(&placeholder_path);
    eprintln!(
        "[mmshell][native-drag-m1] prepared placeholder: {}",
        placeholder_path.display()
    );

    Ok(NativeDragOutResult {
        ok: true,
        cache_path: placeholder_path.to_string_lossy().to_string(),
        message: "M1 占位文件已创建，可先手动拖出验证系统链路。".to_string(),
    })
}

/// 在资源管理器中显示给定路径。 / Reveal given path in Windows Explorer.
#[tauri::command]
fn reveal_in_explorer(path: String) -> Result<(), String> {
    let p = PathBuf::from(&path);
    if !p.exists() {
        return Err(format!("目标路径不存在: {path}"));
    }
    open_in_explorer(&p);
    eprintln!("[mmshell][native-drag-m2] reveal path: {}", p.display());
    Ok(())
}

/// 接收前端调试日志并输出到后端控制台。 / Bridge frontend debug logs to backend console.
#[tauri::command]
fn debug_log(line: String) {
    eprintln!("[mmshell][frontend] {line}");
}

/// 对本地缓存文件执行原生 DoDragDrop。 / Execute native DoDragDrop with cached local file.
#[tauri::command]
fn native_drag_drop_cached(path: String) -> Result<NativeDragDropResult, String> {
    let src = PathBuf::from(&path);
    if !src.exists() {
        return Err(format!("拖拽源不存在: {}", src.display()));
    }
    let cpath = std::ffi::CString::new(path.clone())
        .map_err(|e| format!("构造路径参数失败: {e}"))?;
    let mut effect: u32 = 0;
    let hr = unsafe { mmshell_start_hdrop_drag_from_file_utf8(cpath.as_ptr(), &mut effect as *mut u32) };
    if hr < 0 {
        return Err(format!("C++ hdrop 拖拽失败: HRESULT=0x{:08X}", hr as u32));
    }
    eprintln!("[mmshell][native-drag-cpp] done path={} effect={}", path, effect);
    Ok(NativeDragDropResult { effect })
}

/// 使用 C++ 虚拟拖拽实现从本地缓存文件发起拖拽。 / Start virtual drag from local cache file using C++ implementation.
#[tauri::command]
fn native_drag_drop_cached_virtual(payload: NativeCachedVirtualDragPayload) -> Result<NativeDragDropResult, String> {
    let local = PathBuf::from(&payload.path);
    if !local.exists() {
        return Err(format!("拖拽源不存在: {}", local.display()));
    }
    let cpath = std::ffi::CString::new(payload.path.clone())
        .map_err(|e| format!("构造路径参数失败: {e}"))?;
    let cname = std::ffi::CString::new(payload.display_name.clone())
        .map_err(|e| format!("构造显示名参数失败: {e}"))?;
    let mut effect: u32 = 0;
    let hr = unsafe { mmshell_start_virtual_drag_from_file_utf8(cpath.as_ptr(), cname.as_ptr(), &mut effect as *mut u32) };
    if hr < 0 {
        return Err(format!("C++ 虚拟拖拽失败: HRESULT=0x{:08X}", hr as u32));
    }
    Ok(NativeDragDropResult { effect })
}

/// 执行虚拟文件拖拽（FileDescriptor/FileContents）。 / Execute virtual file drag using FileDescriptor/FileContents.
#[tauri::command]
fn native_drag_drop_virtual(app: AppHandle, payload: NativeVirtualDragPayload) -> Result<NativeDragDropResult, String> {
    let src = PathBuf::from(&payload.source_path);
    let (writer, command_lock) = if let Some(session_id) = payload.sftp_session_id.clone() {
        let store = app.state::<SftpSessionStore>();
        let lock = store
            .sessions
            .lock()
            .map_err(|_| "sftp session store lock poisoned".to_string())?;
        if let Some(s) = lock.get(&session_id) {
            (Some(Arc::clone(&s.writer)), Some(Arc::clone(&s.command_lock)))
        } else {
            (None, None)
        }
    } else {
        (None, None)
    };
    if !src.exists() && writer.is_none() {
        return Err(format!("拖拽源不存在且无可用会话: {}", src.display()));
    }
    let safe_display_name = sanitize_placeholder_name(&payload.display_name);
    let effective_source_path = if payload.source_path.trim().is_empty() {
        let cache_root = resolve_drag_cache_dir()?;
        std::fs::create_dir_all(&cache_root)
            .map_err(|e| format!("创建虚拟拖拽缓存目录失败 {}: {e}", cache_root.display()))?;
        let ts = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|d| d.as_millis())
            .unwrap_or_default();
        cache_root
            .join(format!("virtual-{ts}-{}", safe_display_name))
            .to_string_lossy()
            .to_string()
    } else {
        payload.source_path.clone()
    };

    unsafe { OleInitialize(None).map_err(|e| format!("OleInitialize 失败: {e}"))? };
    let drag_obj = NativeVirtualFileDragObject {
        source_path: effective_source_path.clone(),
        display_name: safe_display_name.clone(),
        sftp_writer: writer,
        sftp_command_lock: command_lock,
        remote_name: payload.remote_name.clone(),
    };
    let data_obj: IDataObject = drag_obj.into();
    let drop_source: IDropSource = data_obj
        .cast()
        .map_err(|e| format!("创建 IDropSource 失败: {e}"))?;
    let mut effect = DROPEFFECT(0);
    let hr = unsafe { DoDragDrop(&data_obj, &drop_source, DROPEFFECT_COPY, &mut effect) };
    unsafe { OleUninitialize() };
    hr.ok().map_err(|e| format!("DoDragDrop(virtual) 失败: {e}"))?;
    eprintln!(
        "[mmshell][native-drag-virtual] done source={} display={} effect={}",
        src.display(),
        safe_display_name,
        effect.0
    );
    Ok(NativeDragDropResult { effect: effect.0 })
}

/// SFTP 原生拖拽主入口：下载到临时目录后发起 DoDragDrop。 / Native SFTP drag entry: download to temp then start DoDragDrop.
#[tauri::command]
fn native_drag_drop_sftp_hdrop(app: AppHandle, payload: NativeSftpHdropPayload) -> Result<NativeDragDropResult, String> {
    let (writer, command_lock) = {
        let store = app.state::<SftpSessionStore>();
        let lock = store
            .sessions
            .lock()
            .map_err(|_| "sftp session store lock poisoned".to_string())?;
        let session = lock
            .get(&payload.sftp_session_id)
            .ok_or_else(|| format!("未找到 SFTP 会话: {}", payload.sftp_session_id))?;
        (Arc::clone(&session.writer), Arc::clone(&session.command_lock))
    };
    let local_path = download_sftp_to_temp_path(
        &writer,
        &command_lock,
        &payload.remote_path,
        &payload.display_name,
        payload.is_dir,
    )?;
    let path_for_hdrop = local_path.to_string_lossy().to_string();
    let cpath = std::ffi::CString::new(path_for_hdrop.clone())
        .map_err(|e| format!("构造路径参数失败: {e}"))?;
    let mut effect: u32 = 0;
    let hr = unsafe { mmshell_start_hdrop_drag_from_file_utf8(cpath.as_ptr(), &mut effect as *mut u32) };
    if hr < 0 {
        return Err(format!("C++ hdrop 拖拽失败: HRESULT=0x{:08X}", hr as u32));
    }
    // 不要立即删除源文件：目标端可能仍在复制，过早删除会触发 Windows 重试/取消弹窗。
    schedule_cleanup_drag_cache_path(local_path.clone(), 120_000);
    eprintln!(
        "[mmshell][native-drag-hdrop] done remote={} local={} effect={}",
        payload.remote_path,
        path_for_hdrop,
        effect
    );
    Ok(NativeDragDropResult { effect })
}

/// 以“松手后下载 + 分块流式供数”方式发起 SFTP 虚拟拖拽。 / Start SFTP virtual drag with post-drop background download and chunked streaming.
#[tauri::command]
fn native_drag_drop_sftp_virtual_streaming(
    app: AppHandle,
    payload: NativeSftpVirtualStreamingPayload,
) -> Result<NativeDragDropResult, String> {
    let (writer, command_lock) = {
        let store = app.state::<SftpSessionStore>();
        let lock = store
            .sessions
            .lock()
            .map_err(|_| "sftp session store lock poisoned".to_string())?;
        let session = lock
            .get(&payload.sftp_session_id)
            .ok_or_else(|| format!("未找到 SFTP 会话: {}", payload.sftp_session_id))?;
        (Arc::clone(&session.writer), Arc::clone(&session.command_lock))
    };

    let base_name = sanitize_placeholder_name(&payload.display_name);
    let cache_root = resolve_drag_cache_dir()?;
    std::fs::create_dir_all(&cache_root)
        .map_err(|e| format!("创建缓存目录失败 {}: {e}", cache_root.display()))?;
    let ts = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_millis())
        .unwrap_or_default();
    let op_dir = cache_root.join(format!("stream-{ts}"));
    std::fs::create_dir_all(&op_dir)
        .map_err(|e| format!("创建流式缓存目录失败 {}: {e}", op_dir.display()))?;
    let local_path = op_dir.join(&base_name);
    let done_marker = op_dir.join(format!("{}.done", base_name));
    let err_marker = op_dir.join(format!("{}.err", base_name));

    spawn_sftp_background_download(
        writer,
        command_lock,
        payload.remote_path.clone(),
        local_path.clone(),
        done_marker.clone(),
        err_marker.clone(),
    );

    let cpath = std::ffi::CString::new(local_path.to_string_lossy().to_string())
        .map_err(|e| format!("构造路径参数失败: {e}"))?;
    let cname = std::ffi::CString::new(payload.display_name.clone())
        .map_err(|e| format!("构造显示名参数失败: {e}"))?;
    let cdone = std::ffi::CString::new(done_marker.to_string_lossy().to_string())
        .map_err(|e| format!("构造完成标记参数失败: {e}"))?;
    let cerr = std::ffi::CString::new(err_marker.to_string_lossy().to_string())
        .map_err(|e| format!("构造错误标记参数失败: {e}"))?;

    let mut effect: u32 = 0;
    let hr = unsafe {
        mmshell_start_virtual_drag_streaming_utf8(
            cpath.as_ptr(),
            cname.as_ptr(),
            cdone.as_ptr(),
            cerr.as_ptr(),
            300_000,
            &mut effect as *mut u32,
        )
    };
    if hr < 0 {
        return Err(format!("C++ streaming virtual 拖拽失败: HRESULT=0x{:08X}", hr as u32));
    }

    schedule_cleanup_drag_cache_path(op_dir.clone(), 180_000);
    Ok(NativeDragDropResult { effect })
}

/// 仅准备 SFTP 拖拽缓存，不立即发起拖拽。 / Prepare SFTP drag cache only, without starting drag immediately.
#[tauri::command]
fn prepare_sftp_drag_cache(app: AppHandle, payload: NativeSftpHdropPayload) -> Result<String, String> {
    let (writer, command_lock) = {
        let store = app.state::<SftpSessionStore>();
        let lock = store
            .sessions
            .lock()
            .map_err(|_| "sftp session store lock poisoned".to_string())?;
        let session = lock
            .get(&payload.sftp_session_id)
            .ok_or_else(|| format!("未找到 SFTP 会话: {}", payload.sftp_session_id))?;
        (Arc::clone(&session.writer), Arc::clone(&session.command_lock))
    };
    let local_path = download_sftp_to_temp_path(
        &writer,
        &command_lock,
        &payload.remote_path,
        &payload.display_name,
        payload.is_dir,
    )?;
    eprintln!(
        "[mmshell][native-drag-hdrop] cache prepared remote={} local={}",
        payload.remote_path,
        local_path.display()
    );
    Ok(local_path.to_string_lossy().to_string())
}

/// 方案B增强：等待鼠标松手并探测 Explorer/桌面目标目录，然后直接下载到目标目录。 / Wait mouse release, detect target folder, then download there directly.
#[tauri::command]
fn native_pick_drop_target_and_download(
    app: AppHandle,
    payload: NativeSftpDropPickPayload,
) -> Result<String, String> {
    // 先发起一个极小占位文件的系统原生拖拽，获得窗口内外一致的系统拖拽光标。
    let ts = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_millis())
        .unwrap_or_default();
    let probe_name = format!("mmshell-drop-probe-{ts}.tmp");
    let probe_path = resolve_drag_cache_dir()?.join(&probe_name);
    std::fs::write(&probe_path, [0u8])
        .map_err(|e| format!("创建拖拽占位文件失败 {}: {e}", probe_path.display()))?;
    let cpath = std::ffi::CString::new(probe_path.to_string_lossy().to_string())
        .map_err(|e| format!("构造占位路径参数失败: {e}"))?;
    let cname = std::ffi::CString::new(payload.display_name.clone())
        .map_err(|e| format!("构造显示名称参数失败: {e}"))?;
    let mut drag_effect: u32 = 0;
    let drag_hr = unsafe { mmshell_start_virtual_drag_from_file_utf8(cpath.as_ptr(), cname.as_ptr(), &mut drag_effect as *mut u32) };
    let _ = std::fs::remove_file(&probe_path);
    if drag_hr < 0 {
        return Err(format!("占位原生拖拽失败: HRESULT=0x{:08X}", drag_hr as u32));
    }
    if drag_effect == 0 {
        return Err("拖拽已取消".to_string());
    }

    // 拖拽结束后再探测目标目录（Explorer/桌面）。
    let mut path_buf = vec![0i8; 4096];
    let mut effect: u32 = 0;
    let hr = unsafe {
        mmshell_detect_drop_target_utf8(
            path_buf.as_mut_ptr(),
            path_buf.len() as u32,
            &mut effect as *mut u32,
        )
    };
    if hr < 0 {
        return Err(format!("探测松手目标目录失败: HRESULT=0x{:08X}", hr as u32));
    }
    if effect == 0 {
        return Err("未识别到可用目标目录".to_string());
    }
    let zero = path_buf.iter().position(|&c| c == 0).unwrap_or(path_buf.len());
    let target_dir = std::str::from_utf8(
        &path_buf[..zero]
            .iter()
            .map(|c| *c as u8)
            .collect::<Vec<u8>>(),
    )
    .map_err(|e| format!("目标目录编码失败: {e}"))?
    .to_string();
    if target_dir.trim().is_empty() {
        return Err("未识别到可用目标目录".to_string());
    }
    let probe_in_target = PathBuf::from(&target_dir).join(&probe_name);
    if probe_in_target.exists() {
        let _ = std::fs::remove_file(&probe_in_target);
    }

    let (writer, command_lock) = {
        let store = app.state::<SftpSessionStore>();
        let lock = store
            .sessions
            .lock()
            .map_err(|_| "sftp session store lock poisoned".to_string())?;
        let session = lock
            .get(&payload.sftp_session_id)
            .ok_or_else(|| format!("未找到 SFTP 会话: {}", payload.sftp_session_id))?;
        (Arc::clone(&session.writer), Arc::clone(&session.command_lock))
    };

    let target_dir_path = PathBuf::from(&target_dir);
    std::fs::create_dir_all(&target_dir_path)
        .map_err(|e| format!("创建目标目录失败 {}: {e}", target_dir_path.display()))?;
    let file_name = sanitize_placeholder_name(&payload.display_name);
    let local_target = target_dir_path.join(&file_name);
    let cmd = if payload.is_dir {
        format!(
            "get -r {} {}\r",
            quote_sftp_arg(&payload.remote_path),
            quote_sftp_arg(&target_dir)
        )
    } else {
        format!(
            "get {} {}\r",
            quote_sftp_arg(&payload.remote_path),
            quote_sftp_arg(&local_target.to_string_lossy())
        )
    };
    let expected = if payload.is_dir {
        target_dir_path.join(file_name)
    } else {
        local_target.clone()
    };
    let _cmd_guard = command_lock
        .lock()
        .map_err(|_| "sftp command lock poisoned".to_string())?;
    {
        let mut lock = writer
            .lock()
            .map_err(|_| "SFTP writer lock poisoned".to_string())?;
        lock.write_all(cmd.as_bytes())
            .map_err(|e| format!("发送 SFTP 下载命令失败: {e}"))?;
        lock.flush().map_err(|e| format!("刷新 SFTP writer 失败: {e}"))?;
    }
    if !wait_for_file_ready(&expected, if payload.is_dir { 240_000 } else { 180_000 }) {
        return Err(format!("下载到目标目录超时: {}", expected.display()));
    }
    if expected.is_file() {
        let _ = wait_for_file_stable(&expected, 30_000, 1200);
    }
    eprintln!(
        "[mmshell][pick-drop] remote={} target={} drag_effect={} detect_effect={}",
        payload.remote_path,
        expected.display(),
        drag_effect,
        effect
    );
    Ok(expected.to_string_lossy().to_string())
}

/// 使用本地占位文件发起原生拖拽（调试用途）。 / Start native drag with local placeholder file (debug purpose).
#[tauri::command]
fn native_drag_drop_placeholder(payload: NativePlaceholderDragPayload) -> Result<NativeDragDropResult, String> {
    let mut base_name = sanitize_placeholder_name(&payload.display_name);
    if !base_name.contains('.') {
        base_name.push_str(".bin");
    }
    let temp_dir = std::env::temp_dir().join("mmshell_native_drag_placeholder");
    std::fs::create_dir_all(&temp_dir)
        .map_err(|e| format!("创建临时目录失败 {}: {e}", temp_dir.display()))?;
    let local_path = temp_dir.join(base_name);
    // 使用小的二进制占位内容，避免文本编辑器把拖拽误当成“文本插入”。
    let binary_placeholder: [u8; 16] = [0x89, 0x42, 0x49, 0x4E, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11];
    std::fs::write(&local_path, binary_placeholder)
        .map_err(|e| format!("写入占位文件失败 {}: {e}", local_path.display()))?;
    let path_for_hdrop = local_path.to_string_lossy().to_string();
    unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED) }
        .ok()
        .map_err(|e| format!("CoInitializeEx 失败: {e}"))?;
    let drag_obj = NativeDragObject {
        path: path_for_hdrop.clone(),
    };
    let data_obj: IDataObject = drag_obj.into();
    let drop_source: IDropSource = data_obj
        .cast()
        .map_err(|e| format!("创建 IDropSource 失败: {e}"))?;
    let mut effect = DROPEFFECT(0);
    let hr = unsafe { DoDragDrop(&data_obj, &drop_source, DROPEFFECT_COPY, &mut effect) };
    unsafe { CoUninitialize() };
    hr.ok().map_err(|e| format!("DoDragDrop(placeholder) 失败: {e}"))?;
    eprintln!(
        "[mmshell][native-drag-placeholder] display={} local={} effect={}",
        payload.display_name,
        path_for_hdrop,
        effect.0
    );
    Ok(NativeDragDropResult { effect: effect.0 })
}

/// 将缓存文件/目录复制到用户目标目录。 / Copy cached file/directory into target directory.
#[tauri::command]
fn copy_cached_to_directory(payload: CopyCachedPayload) -> Result<String, String> {
    let src = PathBuf::from(&payload.source_path);
    if !src.exists() {
        return Err(format!("缓存源不存在: {}", src.display()));
    }
    let target_dir = PathBuf::from(&payload.target_dir);
    if !target_dir.exists() {
        std::fs::create_dir_all(&target_dir)
            .map_err(|e| format!("创建目标目录失败 {}: {e}", target_dir.display()))?;
    }
    let src_name = src
        .file_name()
        .ok_or_else(|| format!("无效源路径: {}", src.display()))?;
    let final_target = target_dir.join(src_name);
    if payload.is_dir {
        copy_dir_recursive(&src, &final_target)?;
    } else {
        std::fs::copy(&src, &final_target)
            .map_err(|e| format!("复制文件失败 {} -> {}: {e}", src.display(), final_target.display()))?;
    }
    eprintln!(
        "[mmshell][native-drag-m4] copied {} -> {}",
        src.display(),
        final_target.display()
    );
    Ok(final_target.to_string_lossy().to_string())
}

/// 启动 Tauri 应用并注册命令与插件。 / Boot Tauri app and register commands/plugins.
#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_fs::init())
        .plugin(tauri_plugin_dialog::init())
        .setup(|_app| {
            ensure_mmshell_config_file();
            eprintln!("[mmshell] 打开了软件 / App opened");
            Ok(())
        })
        .manage(SshSessionStore::default())
        .manage(SftpSessionStore::default())
        .invoke_handler(tauri::generate_handler![
            connect_ssh,
            send_ssh_input,
            resize_ssh,
            disconnect_ssh,
            connect_sftp,
            send_sftp_input,
            disconnect_sftp,
            native_drag_out_begin,
            reveal_in_explorer,
            debug_log,
            copy_cached_to_directory,
            native_drag_drop_cached,
            native_drag_drop_cached_virtual,
            native_drag_drop_virtual,
            native_drag_drop_sftp_virtual_streaming,
            native_pick_drop_target_and_download,
            native_drag_drop_sftp_hdrop,
            prepare_sftp_drag_cache,
            native_drag_drop_placeholder
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}

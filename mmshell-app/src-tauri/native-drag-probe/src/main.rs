use std::mem::ManuallyDrop;
use windows::{
    core::Interface,
    Win32::{
        Foundation::{
            DRAGDROP_S_CANCEL, DRAGDROP_S_DROP, DRAGDROP_S_USEDEFAULTCURSORS, DV_E_FORMATETC,
            E_NOTIMPL, E_UNEXPECTED, HGLOBAL, S_OK,
        },
        System::{
            Com::{IDataObject, IDataObject_Impl, FORMATETC, STGMEDIUM, TYMED_HGLOBAL},
            Memory::{GlobalAlloc, GlobalLock, GlobalUnlock, GMEM_MOVEABLE, GMEM_ZEROINIT},
            Ole::{
                CF_HDROP, DoDragDrop, DROPEFFECT, DROPEFFECT_COPY, IDropSource, IDropSource_Impl,
                OleInitialize, OleUninitialize,
            },
            SystemServices::MODIFIERKEYS_FLAGS,
        },
        UI::Shell::DROPFILES,
    },
};
use windows_core::{implement, BOOL, HRESULT, Result as WinResult};

fn build_hdrop_for_path(path: &str) -> anyhow::Result<HGLOBAL> {
    let mut wide: Vec<u16> = path.encode_utf16().collect();
    wide.push(0);
    wide.push(0);
    let header_size = std::mem::size_of::<DROPFILES>();
    let payload_size = wide.len() * std::mem::size_of::<u16>();
    let total_size = header_size + payload_size;

    let hglobal = unsafe { GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, total_size)? };
    let locked = unsafe { GlobalLock(hglobal) } as *mut u8;
    anyhow::ensure!(!locked.is_null(), "GlobalLock failed");
    unsafe {
        let dropfiles = locked as *mut DROPFILES;
        (*dropfiles).pFiles = header_size as u32;
        (*dropfiles).fNC = BOOL(0);
        (*dropfiles).fWide = BOOL(1);
        let dst = locked.add(header_size) as *mut u16;
        std::ptr::copy_nonoverlapping(wide.as_ptr(), dst, wide.len());
        let _ = GlobalUnlock(hglobal);
    }
    Ok(hglobal)
}

#[implement(IDataObject, IDropSource)]
struct NativeDragObject {
    path: String,
}

impl IDataObject_Impl for NativeDragObject_Impl {
    fn GetData(&self, pformatetcin: *const FORMATETC) -> WinResult<STGMEDIUM> {
        if pformatetcin.is_null() {
            return Err(windows_core::Error::new(E_UNEXPECTED, "FORMATETC is null"));
        }
        let fmt = unsafe { *pformatetcin };
        if fmt.cfFormat != CF_HDROP.0 {
            return Err(windows_core::Error::new(DV_E_FORMATETC, "unsupported format"));
        }
        if fmt.tymed != TYMED_HGLOBAL.0 as u32 {
            return Err(windows_core::Error::new(DV_E_FORMATETC, "unsupported tymed"));
        }
        let hglobal = build_hdrop_for_path(&self.path)
            .map_err(|e| windows_core::Error::new(E_UNEXPECTED, e.to_string()))?;
        let mut medium = STGMEDIUM::default();
        medium.tymed = TYMED_HGLOBAL.0 as u32;
        unsafe {
            medium.u.hGlobal = hglobal;
        }
        medium.pUnkForRelease = ManuallyDrop::new(None);
        Ok(medium)
    }

    fn GetDataHere(&self, _pformatetc: *const FORMATETC, _pmedium: *mut STGMEDIUM) -> WinResult<()> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
    fn QueryGetData(&self, pformatetc: *const FORMATETC) -> HRESULT {
        if pformatetc.is_null() {
            return DV_E_FORMATETC;
        }
        let fmt = unsafe { *pformatetc };
        if fmt.cfFormat == CF_HDROP.0 && fmt.tymed == TYMED_HGLOBAL.0 as u32 {
            S_OK
        } else {
            DV_E_FORMATETC
        }
    }
    fn GetCanonicalFormatEtc(&self, _a: *const FORMATETC, _b: *mut FORMATETC) -> HRESULT { E_NOTIMPL }
    fn SetData(&self, _a: *const FORMATETC, _b: *const STGMEDIUM, _c: BOOL) -> WinResult<()> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
    fn EnumFormatEtc(&self, _dwdirection: u32) -> WinResult<windows::Win32::System::Com::IEnumFORMATETC> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
    fn DAdvise(
        &self,
        _a: *const FORMATETC,
        _b: u32,
        _c: windows_core::Ref<'_, windows::Win32::System::Com::IAdviseSink>,
    ) -> WinResult<u32> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
    fn DUnadvise(&self, _dwconnection: u32) -> WinResult<()> { Err(windows_core::Error::from(E_NOTIMPL)) }
    fn EnumDAdvise(&self) -> WinResult<windows::Win32::System::Com::IEnumSTATDATA> {
        Err(windows_core::Error::from(E_NOTIMPL))
    }
}

impl IDropSource_Impl for NativeDragObject_Impl {
    fn QueryContinueDrag(&self, fescapepressed: BOOL, grfkeystate: MODIFIERKEYS_FLAGS) -> HRESULT {
        if fescapepressed.as_bool() {
            return DRAGDROP_S_CANCEL;
        }
        if (grfkeystate.0 & 0x0001) == 0 {
            return DRAGDROP_S_DROP;
        }
        S_OK
    }
    fn GiveFeedback(&self, _dweffect: DROPEFFECT) -> HRESULT {
        DRAGDROP_S_USEDEFAULTCURSORS
    }
}

fn main() {
    println!("[native-drag-probe] start");
    if let Err(e) = run_probe() {
        eprintln!("[native-drag-probe] failed: {e}");
        std::process::exit(1);
    }
}

fn run_probe() -> anyhow::Result<()> {
    unsafe { OleInitialize(None)? };
    println!("[native-drag-probe] OleInitialize ok");

    let source = std::env::args().nth(1).unwrap_or_else(|| "D:\\MMShell0414\\drag-cache\\probe.txt".to_string());
    if !std::path::Path::new(&source).exists() {
        std::fs::create_dir_all("D:\\MMShell0414\\drag-cache")?;
        std::fs::write(&source, "native drag probe file")?;
    }
    println!("[native-drag-probe] source={source}");
    println!("[native-drag-probe] 开始拖拽，请把文件拖到资源管理器目录后松手");

    let drag_obj = NativeDragObject { path: source.clone() };
    let data_obj: IDataObject = drag_obj.into();
    let drop_source: IDropSource = data_obj.cast()?;
    let mut effect = DROPEFFECT(0);
    let hr = unsafe { DoDragDrop(&data_obj, &drop_source, DROPEFFECT_COPY, &mut effect) };
    println!("[native-drag-probe] DoDragDrop hr={:?} effect={}", hr, effect.0);

    unsafe { OleUninitialize() };
    println!("[native-drag-probe] OleUninitialize ok");
    Ok(())
}

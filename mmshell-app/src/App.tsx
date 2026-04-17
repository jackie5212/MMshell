import { useEffect, useMemo, useRef, useState } from "react";
import { invoke, isTauri } from "@tauri-apps/api/core";
import { listen, type UnlistenFn } from "@tauri-apps/api/event";
import { writeTextFile, readTextFile, exists, mkdir } from "@tauri-apps/plugin-fs";
import { getCurrentWindow } from "@tauri-apps/api/window";

// 扩展Window类型以支持Tauri
declare global {
  interface Window {
    __TAURI__?: {
      invoke?: typeof invoke;
    };
  }
}
import {
  Archive,
  AudioLines,
  Database,
  File,
  FileCode2,
  FileCog,
  FileJson2,
  FileSpreadsheet,
  FileText,
  FileType2,
  FileVideo2,
  Folder,
  FolderPlus,
  Image,
  Presentation,
  ScrollText,
  TerminalSquare,
  Trash2,
  ChevronLeft,
  Star,
  FolderTree,
  Palette,
  Minus,
  Maximize2,
  X,
  FilePlus2,
} from "lucide-react";
import { Terminal } from "xterm";
import { FitAddon } from "xterm-addon-fit";
import "xterm/css/xterm.css";
import "./App.css";

console.log("[MMSHELL] 前端模块已加载 / Frontend module loaded");

type PtyOutputEvent = {
  sessionId: string;
  data: string;
  stream: string;
};

type NativeDragStage = "idle" | "preparing" | "transferring" | "ready" | "failed";

type SftpEntry = {
  name: string;
  isDir: boolean;
};

type Session = {
  id: string;
  name: string;
  host: string;
  port: number;
  user: string;
  group: string;
  password?: string;
};

type TerminalTheme = {
  background: string;
  foreground: string;
  cursor: string;
  cursorAccent: string;
  selectionBackground: string;
  black: string;
  red: string;
  green: string;
  yellow: string;
  blue: string;
  magenta: string;
  cyan: string;
  white: string;
  brightBlack: string;
  brightRed: string;
  brightGreen: string;
  brightYellow: string;
  brightBlue: string;
  brightMagenta: string;
  brightCyan: string;
  brightWhite: string;
};

type ThemePresetId =
  | "shell-dark"
  | "linux-classic"
  | "one-dark"
  | "solarized-dark"
  | "monokai"
  | "custom";
type AppLanguage = "zh-CN" | "en-US";

type I18nKey =
  | "connecting"
  | "notConnected"
  | "menuFile"
  | "menuEdit"
  | "menuView"
  | "menuTools"
  | "menuHelp"
  | "newSession"
  | "openSession"
  | "exit"
  | "copy"
  | "paste"
  | "fullscreen"
  | "splitWindow"
  | "options"
  | "sshKeyManager"
  | "about"
  | "helpDocs"
  | "connect"
  | "disconnect"
  | "refresh"
  | "saveSession"
  | "theme"
  | "session"
  | "directory"
  | "sftpDirectory"
  | "currentPath"
  | "newFolderName"
  | "createDir"
  | "createEmptyFile"
  | "sftpContextCreateDir"
  | "sftpContextCreateFile"
  | "inputFolderNamePrompt"
  | "inputFileNamePrompt"
  | "invalidFolderName"
  | "invalidFileName"
  | "syncDirTerminal"
  | "autoLoadHint"
  | "status"
  | "shortcut"
  | "encoding"
  | "terminalType"
  | "settings"
  | "language"
  | "languageTip"
  | "chineseSimplified"
  | "english"
  | "close"
  | "saveFailed"
  | "sftpAuthHint"
  | "delete"
  | "deleteConfirm"
  | "dir"
  | "file"
  | "groupCommon"
  | "groupProduction"
  | "themeTitle"
  | "themePreset"
  | "custom"
  | "resetShellDark"
  | "done"
  | "newSessionTitle"
  | "sessionName"
  | "hostAddress"
  | "port"
  | "username"
  | "password"
  | "group"
  | "cancel"
  | "save"
  | "inputSessionName"
  | "inputHost"
  | "inputUsername"
  | "inputPassword"
  | "inputGroup"
  | "statusConnecting"
  | "statusConnected"
  | "statusConnectFailed"
  | "statusDisconnected"
  | "statusSessionSaved"
  | "statusCurrentSaved"
  | "welcome"
  | "welcomeHint";

/** 会话凭据等信息的本地 JSON 路径（与 Tauri fs scope 一致） */
const MMSHELL_CONFIG_DIR = "D:/MMShell0414";
const SESSIONS_CONFIG_PATH = `${MMSHELL_CONFIG_DIR}/mmshell_config.json`;
const SSH_LIST_CMD = "ls -la --color=never 2>/dev/null || ls -la";
const I18N: Record<AppLanguage, Record<I18nKey, string>> = {
  "zh-CN": {
    connecting: "正在进入",
    notConnected: "未连接",
    menuFile: "文件",
    menuEdit: "编辑",
    menuView: "视图",
    menuTools: "工具",
    menuHelp: "帮助",
    newSession: "新建会话",
    openSession: "打开会话",
    exit: "退出",
    copy: "复制",
    paste: "粘贴",
    fullscreen: "全屏",
    splitWindow: "分割窗口",
    options: "选项",
    sshKeyManager: "SSH密钥管理",
    about: "关于",
    helpDocs: "帮助文档",
    connect: "连接",
    disconnect: "断开",
    refresh: "刷新",
    saveSession: "保存会话",
    theme: "主题",
    session: "会话",
    directory: "目录",
    sftpDirectory: "SFTP目录",
    currentPath: "当前路径",
    newFolderName: "新文件夹名",
    createDir: "新建目录",
    createEmptyFile: "创建空文件",
    sftpContextCreateDir: "创建文件夹",
    sftpContextCreateFile: "创建空文件",
    inputFolderNamePrompt: "请输入文件夹名称",
    inputFileNamePrompt: "请输入文件名称",
    invalidFolderName: "请输入有效的目录名。",
    invalidFileName: "请输入有效的文件名。",
    syncDirTerminal: "目录和终端同步",
    autoLoadHint: "连接成功后会自动加载目录",
    status: "状态",
    shortcut: "快捷键",
    encoding: "编码",
    terminalType: "终端类型",
    settings: "设置",
    language: "语言",
    languageTip: "修改后立即生效，并会保存到本地配置文件。",
    chineseSimplified: "简体中文",
    english: "English",
    close: "关闭",
    saveFailed: "保存失败：请在 Tauri 桌面版中运行",
    sftpAuthHint: "SFTP 使用 SSH 密码自动登录，无需重复输入。",
    delete: "删除",
    deleteConfirm: "确定删除 {kind}「{name}」？",
    dir: "目录",
    file: "文件",
    groupCommon: "常用",
    groupProduction: "生产",
    themeTitle: "终端主题",
    themePreset: "预设主题",
    custom: "自定义",
    resetShellDark: "重置为 Shell Dark",
    done: "完成",
    newSessionTitle: "新建会话",
    sessionName: "会话名称",
    hostAddress: "主机地址",
    port: "端口号",
    username: "用户名",
    password: "密码",
    group: "分组",
    cancel: "取消",
    save: "保存",
    inputSessionName: "请输入会话名称",
    inputHost: "请输入主机地址",
    inputUsername: "请输入用户名",
    inputPassword: "请输入密码",
    inputGroup: "请输入分组",
    statusConnecting: "连接中...",
    statusConnected: "已连接 ({sid})",
    statusConnectFailed: "连接失败",
    statusDisconnected: "未连接",
    statusSessionSaved: "会话已保存（用户名、地址、端口、密码已写入配置文件）",
    statusCurrentSaved: "当前连接已保存到配置文件",
    welcome: "欢迎使用 MMShell",
    welcomeHint: "请输入连接地址并点击连接按钮",
  },
  "en-US": {
    connecting: "Connecting",
    notConnected: "Disconnected",
    menuFile: "File",
    menuEdit: "Edit",
    menuView: "View",
    menuTools: "Tools",
    menuHelp: "Help",
    newSession: "New Session",
    openSession: "Open Session",
    exit: "Exit",
    copy: "Copy",
    paste: "Paste",
    fullscreen: "Fullscreen",
    splitWindow: "Split Window",
    options: "Options",
    sshKeyManager: "SSH Key Manager",
    about: "About",
    helpDocs: "Help Docs",
    connect: "Connect",
    disconnect: "Disconnect",
    refresh: "Refresh",
    saveSession: "Save Session",
    theme: "Theme",
    session: "Session",
    directory: "Directory",
    sftpDirectory: "SFTP Directory",
    currentPath: "Current Path",
    newFolderName: "New Folder Name",
    createDir: "Create Folder",
    createEmptyFile: "Create Empty File",
    sftpContextCreateDir: "Create Folder",
    sftpContextCreateFile: "Create Empty File",
    inputFolderNamePrompt: "Please enter folder name",
    inputFileNamePrompt: "Please enter file name",
    invalidFolderName: "Please enter a valid directory name.",
    invalidFileName: "Please enter a valid file name.",
    syncDirTerminal: "Sync directory with terminal",
    autoLoadHint: "Directory will load automatically after connection.",
    status: "Status",
    shortcut: "Shortcuts",
    encoding: "Encoding",
    terminalType: "Terminal",
    settings: "Settings",
    language: "Language",
    languageTip: "Changes take effect immediately and are saved locally.",
    chineseSimplified: "Simplified Chinese",
    english: "English",
    close: "Close",
    saveFailed: "Save failed: run in Tauri desktop mode",
    sftpAuthHint: "SFTP logs in with SSH password automatically.",
    delete: "Delete",
    deleteConfirm: "Delete {kind} \"{name}\"?",
    dir: "directory",
    file: "file",
    groupCommon: "Common",
    groupProduction: "Production",
    themeTitle: "Terminal Theme",
    themePreset: "Preset Theme",
    custom: "Custom",
    resetShellDark: "Reset to Shell Dark",
    done: "Done",
    newSessionTitle: "New Session",
    sessionName: "Session Name",
    hostAddress: "Host Address",
    port: "Port",
    username: "Username",
    password: "Password",
    group: "Group",
    cancel: "Cancel",
    save: "Save",
    inputSessionName: "Enter session name",
    inputHost: "Enter host address",
    inputUsername: "Enter username",
    inputPassword: "Enter password",
    inputGroup: "Enter group name",
    statusConnecting: "Connecting...",
    statusConnected: "Connected ({sid})",
    statusConnectFailed: "Connection failed",
    statusDisconnected: "Disconnected",
    statusSessionSaved: "Session saved (username, host, port, password).",
    statusCurrentSaved: "Current connection saved.",
    welcome: "Welcome to MMShell",
    welcomeHint: "Enter address and click Connect",
  },
};
const TERMINAL_THEME_PRESETS: Record<Exclude<ThemePresetId, "custom">, { label: string; theme: TerminalTheme }> = {
  "shell-dark": {
    label: "Shell Dark",
    theme: {
      background: "#0b1220",
      foreground: "#d6deeb",
      cursor: "#f8fafc",
      cursorAccent: "#0b1220",
      selectionBackground: "#2a3c57",
      black: "#1b2432",
      red: "#ff6b6b",
      green: "#526048",
      yellow: "#f6c177",
      blue: "#82aaff",
      magenta: "#c792ea",
      cyan: "#89ddff",
      white: "#d6deeb",
      brightBlack: "#5c6773",
      brightRed: "#ff8b8b",
      brightGreen: "#a0e8af",
      brightYellow: "#ffd89a",
      brightBlue: "#9cc4ff",
      brightMagenta: "#ddb3ff",
      brightCyan: "#b2ebff",
      brightWhite: "#ffffff",
    },
  },
  "linux-classic": {
    label: "Linux Classic",
    theme: {
      background: "#000000",
      foreground: "#d0d0d0",
      cursor: "#d0d0d0",
      cursorAccent: "#000000",
      selectionBackground: "#3a3a3a",
      black: "#000000",
      red: "#aa0000",
      green: "#00aa00",
      yellow: "#aa5500",
      blue: "#0000aa",
      magenta: "#aa00aa",
      cyan: "#00aaaa",
      white: "#aaaaaa",
      brightBlack: "#555555",
      brightRed: "#ff5555",
      brightGreen: "#55ff55",
      brightYellow: "#ffff55",
      brightBlue: "#5555ff",
      brightMagenta: "#ff55ff",
      brightCyan: "#55ffff",
      brightWhite: "#ffffff",
    },
  },
  "one-dark": {
    label: "One Dark",
    theme: {
      background: "#282c34",
      foreground: "#abb2bf",
      cursor: "#528bff",
      cursorAccent: "#282c34",
      selectionBackground: "#3e4451",
      black: "#282c34",
      red: "#e06c75",
      green: "#98c379",
      yellow: "#e5c07b",
      blue: "#61afef",
      magenta: "#c678dd",
      cyan: "#56b6c2",
      white: "#dcdfe4",
      brightBlack: "#5c6370",
      brightRed: "#be5046",
      brightGreen: "#98c379",
      brightYellow: "#d19a66",
      brightBlue: "#61afef",
      brightMagenta: "#c678dd",
      brightCyan: "#56b6c2",
      brightWhite: "#ffffff",
    },
  },
  "solarized-dark": {
    label: "Solarized Dark",
    theme: {
      background: "#002b36",
      foreground: "#93a1a1",
      cursor: "#93a1a1",
      cursorAccent: "#002b36",
      selectionBackground: "#073642",
      black: "#073642",
      red: "#dc322f",
      green: "#859900",
      yellow: "#b58900",
      blue: "#268bd2",
      magenta: "#d33682",
      cyan: "#2aa198",
      white: "#eee8d5",
      brightBlack: "#002b36",
      brightRed: "#cb4b16",
      brightGreen: "#586e75",
      brightYellow: "#657b83",
      brightBlue: "#839496",
      brightMagenta: "#6c71c4",
      brightCyan: "#93a1a1",
      brightWhite: "#fdf6e3",
    },
  },
  monokai: {
    label: "Monokai",
    theme: {
      background: "#272822",
      foreground: "#f8f8f2",
      cursor: "#f8f8f0",
      cursorAccent: "#272822",
      selectionBackground: "#49483e",
      black: "#272822",
      red: "#f92672",
      green: "#a6e22e",
      yellow: "#f4bf75",
      blue: "#66d9ef",
      magenta: "#ae81ff",
      cyan: "#2aa198",
      white: "#f8f8f2",
      brightBlack: "#75715e",
      brightRed: "#f92672",
      brightGreen: "#a6e22e",
      brightYellow: "#e6db74",
      brightBlue: "#66d9ef",
      brightMagenta: "#fd5ff0",
      brightCyan: "#a1efe4",
      brightWhite: "#f9f8f5",
    },
  },
};

const CUSTOM_THEME_FIELDS: Array<{
  key: keyof TerminalTheme;
  label: Record<AppLanguage, string>;
  description: Record<AppLanguage, string>;
}> = [
  { key: "background", label: { "zh-CN": "背景", "en-US": "Background" }, description: { "zh-CN": "终端整体底色。", "en-US": "Terminal background color." } },
  { key: "foreground", label: { "zh-CN": "前景文字", "en-US": "Foreground" }, description: { "zh-CN": "普通文本默认颜色。", "en-US": "Default text color." } },
  { key: "cursor", label: { "zh-CN": "光标", "en-US": "Cursor" }, description: { "zh-CN": "输入光标颜色。", "en-US": "Cursor color." } },
  { key: "cursorAccent", label: { "zh-CN": "光标内文字", "en-US": "Cursor Accent" }, description: { "zh-CN": "块状光标覆盖时文字颜色。", "en-US": "Text color inside block cursor." } },
  { key: "selectionBackground", label: { "zh-CN": "选中背景", "en-US": "Selection" }, description: { "zh-CN": "鼠标选中文本的高亮背景。", "en-US": "Selected text background." } },
  { key: "black", label: { "zh-CN": "标准黑", "en-US": "Black" }, description: { "zh-CN": "ANSI 30。", "en-US": "ANSI 30." } },
  { key: "red", label: { "zh-CN": "标准红", "en-US": "Red" }, description: { "zh-CN": "ANSI 31。", "en-US": "ANSI 31." } },
  { key: "green", label: { "zh-CN": "标准绿", "en-US": "Green" }, description: { "zh-CN": "ANSI 32。", "en-US": "ANSI 32." } },
  { key: "yellow", label: { "zh-CN": "标准黄", "en-US": "Yellow" }, description: { "zh-CN": "ANSI 33。", "en-US": "ANSI 33." } },
  { key: "blue", label: { "zh-CN": "标准蓝", "en-US": "Blue" }, description: { "zh-CN": "ANSI 34。", "en-US": "ANSI 34." } },
  { key: "magenta", label: { "zh-CN": "标准洋红", "en-US": "Magenta" }, description: { "zh-CN": "ANSI 35。", "en-US": "ANSI 35." } },
  { key: "cyan", label: { "zh-CN": "标准青", "en-US": "Cyan" }, description: { "zh-CN": "ANSI 36。", "en-US": "ANSI 36." } },
  { key: "white", label: { "zh-CN": "标准白", "en-US": "White" }, description: { "zh-CN": "ANSI 37。", "en-US": "ANSI 37." } },
  { key: "brightBlack", label: { "zh-CN": "高亮黑", "en-US": "Bright Black" }, description: { "zh-CN": "ANSI 90。", "en-US": "ANSI 90." } },
  { key: "brightRed", label: { "zh-CN": "高亮红", "en-US": "Bright Red" }, description: { "zh-CN": "ANSI 91。", "en-US": "ANSI 91." } },
  { key: "brightGreen", label: { "zh-CN": "高亮绿", "en-US": "Bright Green" }, description: { "zh-CN": "ANSI 92。", "en-US": "ANSI 92." } },
  { key: "brightYellow", label: { "zh-CN": "高亮黄", "en-US": "Bright Yellow" }, description: { "zh-CN": "ANSI 93。", "en-US": "ANSI 93." } },
  { key: "brightBlue", label: { "zh-CN": "高亮蓝", "en-US": "Bright Blue" }, description: { "zh-CN": "ANSI 94。", "en-US": "ANSI 94." } },
  { key: "brightMagenta", label: { "zh-CN": "高亮洋红", "en-US": "Bright Magenta" }, description: { "zh-CN": "ANSI 95。", "en-US": "ANSI 95." } },
  { key: "brightCyan", label: { "zh-CN": "高亮青", "en-US": "Bright Cyan" }, description: { "zh-CN": "ANSI 96。", "en-US": "ANSI 96." } },
  { key: "brightWhite", label: { "zh-CN": "高亮白", "en-US": "Bright White" }, description: { "zh-CN": "ANSI 97。", "en-US": "ANSI 97." } },
];

/** 获取文件名后缀（不含点，统一转小写）。 / Get lowercase file extension without dot. */
function getFileExtension(name: string): string {
  const parts = name.toLowerCase().split(".");
  if (parts.length < 2) return "";
  return parts[parts.length - 1];
}

/** 根据文件名返回 SFTP 列表中对应的文件图标组件。 / Return icon component by file name for SFTP list. */
function getSftpIconByName(name: string) {
  const ext = getFileExtension(name);
  if (["ini", "conf", "config", "toml", "yaml", "yml", "env", "properties"].includes(ext)) {
    return FileCog;
  }
  if (["db", "sqlite", "sqlite3", "mdb", "accdb"].includes(ext)) {
    return Database;
  }
  if (["txt", "log", "md", "rtf"].includes(ext)) {
    return FileText;
  }
  if (["json"].includes(ext)) {
    return FileJson2;
  }
  if (["xml", "csv"].includes(ext)) {
    return ScrollText;
  }
  if (
    [
      "rs",
      "ts",
      "tsx",
      "js",
      "jsx",
      "py",
      "go",
      "java",
      "c",
      "cpp",
      "h",
      "hpp",
      "cs",
      "php",
      "rb",
      "swift",
      "kt",
      "sql",
      "sh",
      "bat",
      "ps1",
    ].includes(ext)
  ) {
    return FileCode2;
  }
  if (["exe", "dll", "so", "dylib", "bin"].includes(ext)) {
    return TerminalSquare;
  }
  if (["zip", "rar", "7z", "tar", "gz", "bz2", "xz"].includes(ext)) {
    return Archive;
  }
  if (["png", "jpg", "jpeg", "gif", "webp", "bmp", "svg", "ico"].includes(ext)) {
    return Image;
  }
  if (["mp3", "wav", "flac", "ogg", "aac", "m4a"].includes(ext)) {
    return AudioLines;
  }
  if (["mp4", "mkv", "avi", "mov", "wmv", "webm"].includes(ext)) {
    return FileVideo2;
  }
  if (["xls", "xlsx", "ods"].includes(ext)) {
    return FileSpreadsheet;
  }
  if (["ppt", "pptx", "odp"].includes(ext)) {
    return Presentation;
  }
  if (["pdf", "doc", "docx"].includes(ext)) {
    return FileType2;
  }
  return File;
}

/** 根据文件类型返回 SFTP 列表图标颜色。 / Return SFTP icon color by file type. */
function getSftpIconColorByName(name: string, isDir: boolean): string {
  if (isDir) return "#1d4f91";
  const ext = getFileExtension(name);
  if (["ini", "conf", "config", "toml", "yaml", "yml", "env", "properties"].includes(ext)) {
    return "#c26d16";
  }
  if (["db", "sqlite", "sqlite3", "mdb", "accdb"].includes(ext)) {
    return "#7a3fc2";
  }
  if (["txt", "log", "md", "rtf"].includes(ext)) {
    return "#5b6776";
  }
  if (["json", "xml", "csv"].includes(ext)) {
    return "#0e7a86";
  }
  if (
    [
      "rs",
      "ts",
      "tsx",
      "js",
      "jsx",
      "py",
      "go",
      "java",
      "c",
      "cpp",
      "h",
      "hpp",
      "cs",
      "php",
      "rb",
      "swift",
      "kt",
      "sql",
      "sh",
      "bat",
      "ps1",
    ].includes(ext)
  ) {
    return "#1270b8";
  }
  if (["exe", "dll", "so", "dylib", "bin"].includes(ext)) {
    return "#1d3557";
  }
  if (["zip", "rar", "7z", "tar", "gz", "bz2", "xz"].includes(ext)) {
    return "#8b5e3c";
  }
  if (["png", "jpg", "jpeg", "gif", "webp", "bmp", "svg", "ico"].includes(ext)) {
    return "#258f4a";
  }
  if (["mp3", "wav", "flac", "ogg", "aac", "m4a"].includes(ext)) {
    return "#b23fa1";
  }
  if (["mp4", "mkv", "avi", "mov", "wmv", "webm"].includes(ext)) {
    return "#b53939";
  }
  if (["xls", "xlsx", "ods"].includes(ext)) {
    return "#0f8d66";
  }
  if (["ppt", "pptx", "odp"].includes(ext)) {
    return "#d25522";
  }
  if (["pdf", "doc", "docx"].includes(ext)) {
    return "#9f3e3e";
  }
  return "#4b5c73";
}

/** 清理终端输出中的 ANSI 控制序列。 / Strip ANSI control sequences from terminal output. */
function stripAnsiControls(text: string): string {
  return text
    // CSI / SGR 等：\x1b[ ... letter
    .replace(/\x1B\[[0-?]*[ -/]*[@-~]/g, "")
    // OSC：\x1b] ... \x07
    .replace(/\x1B\][^\x07]*(?:\x07|\x1B\\)/g, "")
    // 其他 C0 控制字符（保留 \r \n \t 便于分行）
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, "");
}

/** 解析 SFTP 传输进度文本（上传/下载）。 / Parse SFTP transfer progress text (upload/download). */
function parseSftpTransferProgress(rawChunk: string): { file: string; percent: number; speed: string } | null {
  const chunk = stripAnsiControls(rawChunk);
  const explicit = chunk.match(/(?:Uploading|Downloading)\s+(.+?)\s+(\d{1,3})%\s+([\d.]+[KMGT]?B\/s)/i);
  if (explicit) {
    return {
      file: explicit[1].trim(),
      percent: Math.min(100, parseInt(explicit[2], 10)),
      speed: explicit[3],
    };
  }
  const genericMatches = Array.from(
    chunk.matchAll(/(?:^|\r?\n)\s*([^\r\n]+?)\s+(\d{1,3})%\s+[\d.]+[KMGT]?B\s+([\d.]+[KMGT]?B\/s)/g)
  );
  const last = genericMatches.length > 0 ? genericMatches[genericMatches.length - 1] : null;
  if (!last) return null;
  return {
    file: last[1].trim(),
    percent: Math.min(100, parseInt(last[2], 10)),
    speed: last[3],
  };
}

/** 从 SFTP 原始输出中解析当前目录文件列表。 / Parse current directory entries from raw SFTP output. */
function parseSftpEntries(raw: string): SftpEntry[] {
  const sanitized = stripAnsiControls(raw);
  const latestLsIndex = Math.max(
    sanitized.lastIndexOf("sftp> ls -la"),
    sanitized.lastIndexOf("sftp> ls -l")
  );
  const parsingSource = latestLsIndex >= 0 ? sanitized.slice(latestLsIndex) : sanitized;
  const lines = parsingSource.split(/\r?\n/);
  const entries: SftpEntry[] = [];
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) continue;
    if (trimmed.startsWith("sftp>")) continue;
    if (/password:/i.test(trimmed)) continue;
    if (trimmed.startsWith("Connected to ")) continue;
    if (trimmed.startsWith("Remote working directory:")) continue;
    if (trimmed.startsWith("Fetching ")) continue;
    if (trimmed.startsWith("usage:")) continue;
    if (trimmed.startsWith("Invalid command")) continue;
    if (trimmed.startsWith("Couldn't")) continue;
    if (trimmed.startsWith("File")) continue;
    if (trimmed.startsWith("remote")) continue;

    if (/^[d\-l]/.test(trimmed)) {
      const parts = trimmed.split(/\s+/);
      if (parts.length >= 9) {
        entries.push({
          name: parts.slice(8).join(" "),
          isDir: parts[0].startsWith("d"),
        });
        continue;
      }
    }
  }
  const map = new Map<string, SftpEntry>();
  for (const entry of entries) {
    if (!map.has(entry.name)) {
      map.set(entry.name, entry);
    }
  }
  return Array.from(map.values()).filter((e) => e.name !== "." && e.name !== "..");
}

/** 对 SFTP 命令参数做安全转义。 / Escape SFTP command arguments safely. */
function quoteSftpArg(name: string): string {
  if (/^[a-zA-Z0-9._@\-/+]+$/.test(name)) return name;
  return `"${name.replace(/\\/g, "\\\\").replace(/"/g, '\\"')}"`;
}

/** 解析工具栏 user@host:port，与 Rust `parse_address` 行为一致。 / Parse user@host:port from toolbar, aligned with Rust parse_address. */
function parseConnectionAddress(addr: string): { user: string; host: string; port: number } | null {
  const trimmed = addr.trim();
  const at = trimmed.lastIndexOf("@");
  if (at <= 0) return null;
  const user = trimmed.slice(0, at).trim();
  let rest = trimmed.slice(at + 1).trim();
  if (!user || !rest) return null;
  let port = 22;
  const colon = rest.lastIndexOf(":");
  if (colon > 0) {
    const p = parseInt(rest.slice(colon + 1), 10);
    if (!Number.isNaN(p) && p > 0 && p <= 65535) {
      port = p;
      rest = rest.slice(0, colon).trim();
    }
  }
  if (!rest) return null;
  return { user, host: rest, port };
}

/** 从 shell 输出中提取最近一次提示符路径。 / Extract latest prompt path from shell output. */
function extractPwdFromPrompt(output: string): string | null {
  // 先清理 ANSI 控制码，避免彩色提示符影响匹配。
  const clean = output
    .replace(/\x1B\[[0-?]*[ -/]*[@-~]/g, "")
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, "");
  // 匹配常见提示符，如：
  // root@luckfox:/#  user@host:/home/user$  root@host:~#
  const regex = /(?:^|\r?\n)\s*[^\s@]+@[^:\s]+:([^\r\n#$>]+)\s*[#$>]/g;
  let match: RegExpExecArray | null = null;
  let lastPath: string | null = null;
  while (true) {
    match = regex.exec(clean);
    if (!match) break;
    const p = match[1]?.trim();
    if (p) lastPath = p;
  }
  if (!lastPath) return null;
  if (lastPath === "~") return "/";
  if (lastPath.startsWith("~/")) return `/${lastPath.slice(2)}`;
  if (!lastPath.startsWith("/")) return null;
  return lastPath;
}

/** 主界面组件：负责连接、会话、SFTP 与主题管理。 / Main UI component for connection, sessions, SFTP, and themes. */
function App() {
  const [address, setAddress] = useState("root@192.168.0.106:22");
  const [password, setPassword] = useState("");
  const [sessionId, setSessionId] = useState<string | null>(null);
  const [sftpSessionId, setSftpSessionId] = useState<string | null>(null);
  const [isConnecting, setIsConnecting] = useState(false);
  const [connectingDotCount, setConnectingDotCount] = useState(1);
  const [status, setStatus] = useState(I18N["zh-CN"].statusDisconnected);
  const [error, setError] = useState("");
  const [appLanguage, setAppLanguage] = useState<AppLanguage>("zh-CN");
  const [sftpEntries, setSftpEntries] = useState<SftpEntry[]>([]);
  const [sftpPath, setSftpPath] = useState("/");
  const [sftpHeaderMenu, setSftpHeaderMenu] = useState<{ visible: boolean; x: number; y: number }>({
    visible: false,
    x: 0,
    y: 0,
  });
  const [sessionMenu, setSessionMenu] = useState<{ visible: boolean; x: number; y: number; session: Session | null }>({
    visible: false,
    x: 0,
    y: 0,
    session: null,
  });

  const [sessionListMenu, setSessionListMenu] = useState<{ visible: boolean; x: number; y: number }>({
    visible: false,
    x: 0,
    y: 0,
  });
  const [syncDirEnabled, setSyncDirEnabled] = useState(true);
  const [sftpProgress, setSftpProgress] = useState<{ file: string; percent: number; speed: string } | null>(null);
  const [nativeDragStatus, setNativeDragStatus] = useState<{
    stage: NativeDragStage;
    item: string;
    targetPath: string;
    message: string;
  }>({
    stage: "idle",
    item: "",
    targetPath: "",
    message: "",
  });
  const [draggingExportName, setDraggingExportName] = useState("");
  const [sessionPanelCollapsed, setSessionPanelCollapsed] = useState(false);
  const [directoryPanelCollapsed, setDirectoryPanelCollapsed] = useState(false);
  const [showNewSessionModal, setShowNewSessionModal] = useState(false);
  const [showThemeModal, setShowThemeModal] = useState(false);
  const [showSettingsModal, setShowSettingsModal] = useState(false);
  const [themePresetId, setThemePresetId] = useState<ThemePresetId>("shell-dark");
  const [customTheme, setCustomTheme] = useState<TerminalTheme>(TERMINAL_THEME_PRESETS["shell-dark"].theme);
  const [newSession, setNewSession] = useState({
    name: "",
    host: "",
    port: 22,
    user: "root",
    password: "",
    group: I18N["zh-CN"].groupCommon
  });
  const [sessions, setSessions] = useState<Session[]>([]);
  const sftpRawRef = useRef("");
  const sftpReadyRef = useRef(false);
  const sftpAutoListedRef = useRef(false);
  const sftpProgressHideTimerRef = useRef<number | null>(null);

  const termHostRef = useRef<HTMLDivElement | null>(null);
  const terminalRef = useRef<Terminal | null>(null);
  const fitRef = useRef<FitAddon | null>(null);
  const sessionIdRef = useRef<string | null>(null);
  const sftpSessionIdRef = useRef<string | null>(null);
  const sshAwaitingPasswordRef = useRef(false);
  const sshPasswordBufferRef = useRef("");
  const sharedPasswordRef = useRef("");
  const passwordRef = useRef("");
  const waitingForShellPromptRef = useRef(false);
  const connectStartedAtRef = useRef(0);
  const connectMinimumVisibleMs = 1200;
  const syncDirEnabledRef = useRef(true);
  const pushDebugLog = (line: string) => {
    if (!isTauri()) return;
    void invoke("debug_log", { line }).catch(() => {});
  };
  const setGlobalDragCursor = (active: boolean) => {
    if (typeof document === "undefined") return;
    if (active) {
      document.body.classList.add("drag-cursor-active");
      document.documentElement.classList.add("drag-cursor-active");
      document.body.style.cursor = "copy";
      document.documentElement.style.cursor = "copy";
      return;
    }
    document.body.classList.remove("drag-cursor-active");
    document.documentElement.classList.remove("drag-cursor-active");
    document.body.style.cursor = "";
    document.documentElement.style.cursor = "";
  };
  const setNativeDragStage = (
    stage: NativeDragStage,
    item: string,
    targetPath: string,
    message: string
  ) => {
    setNativeDragStatus({ stage, item, targetPath, message });
    pushDebugLog(`[M4] stage=${stage} item=${item} target=${targetPath} msg=${message}`);
  };
  const clearSftpProgressHideTimer = () => {
    if (sftpProgressHideTimerRef.current != null) {
      window.clearTimeout(sftpProgressHideTimerRef.current);
      sftpProgressHideTimerRef.current = null;
    }
  };
  const scheduleSftpProgressAutoHide = (delayMs: number) => {
    clearSftpProgressHideTimer();
    sftpProgressHideTimerRef.current = window.setTimeout(() => {
      setSftpProgress(null);
      sftpProgressHideTimerRef.current = null;
    }, delayMs);
  };
  const pushSftpProgress = (next: { file: string; percent: number; speed: string }) => {
    setSftpProgress(next);
    // 若持续没有新进度，自动隐藏，避免“卡住”。
    scheduleSftpProgressAutoHide(next.percent >= 100 ? 1000 : 5000);
  };

  /** 保证“连接中”状态至少展示固定时长，避免闪烁。 / Keep connecting state visible for a minimum duration. */
  function finishConnectingWithMinimumDelay() {
    const elapsed = Date.now() - connectStartedAtRef.current;
    const delay = Math.max(0, connectMinimumVisibleMs - elapsed);
    window.setTimeout(() => setIsConnecting(false), delay);
  }

  /** 当拖拽传输进行中时，阻止其他 SFTP 操作以避免并发冲突。 / Block non-drag SFTP operations while drag transfer is running. */
  function guardSftpBusy(actionName: string): boolean {
    if (!sftpTransferBusyRef.current) return false;
    const msg =
      appLanguage === "zh-CN"
        ? `当前正在执行拖拽传输，请稍后再${actionName}`
        : `Drag transfer in progress. Please retry ${actionName} later.`;
    setError(msg);
    return true;
  }

  const sftpPathRef = useRef("/");
  const lastSyncedPromptPathRef = useRef<string>("/");
  const sshOutputTailRef = useRef("");
  const sftpPasswordSentRef = useRef(false);
  const recentlyCreatedSftpDirsRef = useRef<Map<string, number>>(new Map());
  const nativeDragRunningRef = useRef(false);
  const sftpTransferBusyRef = useRef(false);
  const dragTraceSeqRef = useRef(0);
  const activeTerminalTheme = useMemo<TerminalTheme>(() => {
    if (themePresetId === "custom") return customTheme;
    return TERMINAL_THEME_PRESETS[themePresetId].theme;
  }, [themePresetId, customTheme]);
  const t = useMemo(() => I18N[appLanguage], [appLanguage]);
  /** 读取国际化文案，并替换占位符变量。 / Read i18n text and replace placeholder variables. */
  function tr(key: I18nKey, vars?: Record<string, string>): string {
    let text = t[key];
    if (vars) {
      for (const [k, v] of Object.entries(vars)) {
        text = text.split(`{${k}}`).join(v);
      }
    }
    return text;
  }
  /** 规范化显示分组名，兼容中英文常用分组。 / Normalize group labels for Chinese and English names. */
  function displayGroupName(group: string): string {
    const normalized = group.trim().toLowerCase();
    if (normalized === "常用" || normalized === "common") return t.groupCommon;
    if (normalized === "生产" || normalized === "production") return t.groupProduction;
    return group;
  }
  const tabTitle = useMemo(() => {
    if (isConnecting) return `${t.connecting}${".".repeat(connectingDotCount)}`;
    return address || t.notConnected;
  }, [isConnecting, connectingDotCount, address, t]);

  useEffect(() => {
    console.log("[MMSHELL] 打开了软件 / App opened");
  }, []);

  useEffect(() => {
    if (typeof window !== "undefined" && isTauri()) {
      void (async () => {
        await ensureSessionsConfigFile();
        await loadSessionsFromFile();
      })();
    }
  }, []);

  /** 合并加载的主题字段并补齐默认值。 / Merge loaded theme with defaults for missing keys. */
  function mergeLoadedTerminalTheme(partial: Partial<TerminalTheme> | undefined): TerminalTheme {
    return { ...TERMINAL_THEME_PRESETS["shell-dark"].theme, ...partial };
  }

  /** 判断字符串是否为合法主题预设 ID。 / Check whether value is a valid theme preset id. */
  function isThemePresetId(v: string): v is ThemePresetId {
    return (
      v === "shell-dark" ||
      v === "linux-classic" ||
      v === "one-dark" ||
      v === "solarized-dark" ||
      v === "monokai" ||
      v === "custom"
    );
  }

  /** 启动时确保 D:\\MMShell0414 下存在 mmshell_config.json，用于保存会话名、主机、端口、用户名、密码等。 / Ensure config file exists at startup. */
  async function ensureSessionsConfigFile() {
    try {
      if (typeof window === "undefined" || !isTauri()) return;
      if (!(await exists(MMSHELL_CONFIG_DIR))) {
        await mkdir(MMSHELL_CONFIG_DIR, { recursive: true });
      }
      if (!(await exists(SESSIONS_CONFIG_PATH))) {
        const initial = {
          version: 1,
          sessions: [] as Session[],
          settings: {
            language: "zh-CN" as AppLanguage,
            themePresetId: "shell-dark" as ThemePresetId,
            terminalTheme: TERMINAL_THEME_PRESETS["shell-dark"].theme,
          },
          lastUpdated: new Date().toISOString(),
        };
        await writeTextFile(SESSIONS_CONFIG_PATH, JSON.stringify(initial, null, 2));
        console.log("[Config] 已创建会话配置文件:", SESSIONS_CONFIG_PATH);
      }
    } catch (error) {
      console.error("[Config] 初始化会话配置文件失败:", error);
    }
  }

  useEffect(() => {
    sessionIdRef.current = sessionId;
  }, [sessionId]);

  useEffect(() => {
    passwordRef.current = password;
  }, [password]);

  useEffect(() => {
    syncDirEnabledRef.current = syncDirEnabled;
  }, [syncDirEnabled]);

  useEffect(() => {
    if (!sftpHeaderMenu.visible) return;
    /** 点击页面其他区域时关闭 SFTP 头部右键菜单。 / Close SFTP header context menu when clicking outside. */
    function closeMenu() {
      setSftpHeaderMenu((prev) => ({ ...prev, visible: false }));
    }
    window.addEventListener("click", closeMenu);
    return () => window.removeEventListener("click", closeMenu);
  }, [sftpHeaderMenu.visible]);

  useEffect(() => {
    if (!sessionMenu.visible) return;
    /** 点击页面其他区域时关闭会话右键菜单。 / Close session context menu when clicking outside. */
    function closeMenu() {
      setSessionMenu((prev) => ({ ...prev, visible: false }));
    }
    window.addEventListener("click", closeMenu);
    return () => window.removeEventListener("click", closeMenu);
  }, [sessionMenu.visible]);

  useEffect(() => {
    if (!sessionListMenu.visible) return;
    /** 点击页面其他区域时关闭会话列表右键菜单。 / Close session list context menu when clicking outside. */
    function closeMenu() {
      setSessionListMenu((prev) => ({ ...prev, visible: false }));
    }
    window.addEventListener("click", closeMenu);
    return () => window.removeEventListener("click", closeMenu);
  }, [sessionListMenu.visible]);

  useEffect(() => {
    sftpPathRef.current = sftpPath;
  }, [sftpPath]);

  useEffect(() => {
    if (!isConnecting) {
      setConnectingDotCount(1);
      return;
    }
    const timer = window.setInterval(() => {
      setConnectingDotCount((prev) => (prev >= 3 ? 1 : prev + 1));
    }, 400);
    return () => window.clearInterval(timer);
  }, [isConnecting]);

  useEffect(() => {
    if (!isConnecting) return;
    // 兜底：若远端长时间未出现提示符，最多保留 20 秒“正在进入”提示。
    const timeout = window.setTimeout(() => {
      waitingForShellPromptRef.current = false;
      finishConnectingWithMinimumDelay();
    }, 20000);
    return () => window.clearTimeout(timeout);
  }, [isConnecting]);

  useEffect(() => {
    const term = terminalRef.current;
    if (!term) return;
    term.options.theme = activeTerminalTheme;
  }, [activeTerminalTheme]);

  useEffect(() => {
    sftpSessionIdRef.current = sftpSessionId;
  }, [sftpSessionId]);

  useEffect(() => {
    if (!termHostRef.current || terminalRef.current) return;

    const term = new Terminal({
      cursorBlink: true,
      fontFamily: "Consolas, 'Cascadia Mono', monospace",
      fontSize: 14,
      theme: activeTerminalTheme,
      scrollback: 5000,
      convertEol: true,
    });
    const fitAddon = new FitAddon();
    term.loadAddon(fitAddon);
    term.open(termHostRef.current);
    fitAddon.fit();
    term.focus();

    term.onData((data) => {
      if (sshAwaitingPasswordRef.current) {
        if (data === "\r" || data === "\n") {
          sharedPasswordRef.current = sshPasswordBufferRef.current;
          sshPasswordBufferRef.current = "";
          sshAwaitingPasswordRef.current = false;
          const sftpSid = sftpSessionIdRef.current;
          if (sftpSid && sharedPasswordRef.current && !sftpPasswordSentRef.current) {
            sftpPasswordSentRef.current = true;
            void invoke("send_sftp_input", {
              sessionId: sftpSid,
              input: `${sharedPasswordRef.current}\r`,
            }).catch((err) => setError(String(err)));
          }
        } else if (data === "\u007f" || data === "\b") {
          sshPasswordBufferRef.current = sshPasswordBufferRef.current.slice(0, -1);
        } else if (!data.startsWith("\u001b")) {
          sshPasswordBufferRef.current += data;
        }
      }

      const sid = sessionIdRef.current;
      if (!sid) return;
      void invoke("send_ssh_input", { sessionId: sid, input: data }).catch((err) => {
        setError(String(err));
      });
    });

    term.attachCustomKeyEventHandler((ev) => {
      if (!ev.ctrlKey) return true;
      const key = ev.key.toLowerCase();
      if (key === "v" && ev.type === "keydown") {
        ev.preventDefault();
        const sid = sessionIdRef.current;
        if (!sid) return false;
        void navigator.clipboard.readText().then((text) => {
          if (!text) return;
          void invoke("send_ssh_input", { sessionId: sid, input: text });
        });
        return false;
      }
      if (key === "c" && ev.type === "keydown") {
        const selected = term.getSelection();
        if (selected) {
          ev.preventDefault();
          void navigator.clipboard.writeText(selected);
          return false;
        }
        const sid = sessionIdRef.current;
        if (sid) {
          ev.preventDefault();
          void invoke("send_ssh_input", { sessionId: sid, input: "\u0003" });
          return false;
        }
      }
      return true;
    });

    const syncResize = () => {
      fitAddon.fit();
      const sid = sessionIdRef.current;
      if (!sid) return;
      void invoke("resize_ssh", {
        sessionId: sid,
        cols: term.cols,
        rows: term.rows,
      }).catch(() => {});
    };

    const resizeObserver = new ResizeObserver(syncResize);
    resizeObserver.observe(termHostRef.current);
    window.addEventListener("resize", syncResize);

    terminalRef.current = term;
    fitRef.current = fitAddon;

    return () => {
      resizeObserver.disconnect();
      window.removeEventListener("resize", syncResize);
      term.dispose();
      terminalRef.current = null;
      fitRef.current = null;
    };
  }, []);

  useEffect(() => {
    // 只在Tauri环境中监听事件
    if (typeof window === "undefined" || !isTauri()) return;
    
    let unlisten: UnlistenFn | null = null;
    let disposed = false;
    void (async () => {
      try {
        console.log("[SSH] registering ssh-output listener");
        unlisten = await listen<PtyOutputEvent>("ssh-output", (event) => {
          const payload = event.payload;
          if (payload.sessionId !== sessionIdRef.current) {
            console.log("[SSH] ignore output: sid mismatch", payload.sessionId, sessionIdRef.current);
            return;
          }
          console.log("[SSH] output chunk", {
            sid: payload.sessionId,
            stream: payload.stream,
            size: payload.data.length,
            preview: payload.data.slice(0, 120),
          });
          if (/password:/i.test(payload.data)) {
            sshAwaitingPasswordRef.current = true;
            sshPasswordBufferRef.current = "";
            // 如果已经有密码，自动发送
            if (sharedPasswordRef.current) {
              const sid = sessionIdRef.current;
              if (sid) {
                void invoke("send_ssh_input", {
                  sessionId: sid,
                  input: `${sharedPasswordRef.current}\r`
                }).catch((err) => setError(String(err)));
              }
            }
          }
          const nextTail = `${sshOutputTailRef.current}${payload.data}`.slice(-4096);
          sshOutputTailRef.current = nextTail;
          const promptPath = extractPwdFromPrompt(nextTail);
          if (promptPath && waitingForShellPromptRef.current) {
            waitingForShellPromptRef.current = false;
            finishConnectingWithMinimumDelay();
          }
          if (
            promptPath &&
            syncDirEnabledRef.current &&
            !sftpTransferBusyRef.current &&
            sftpReadyRef.current &&
            sftpSessionIdRef.current &&
            promptPath !== sftpPathRef.current &&
            promptPath !== lastSyncedPromptPathRef.current
          ) {
            lastSyncedPromptPathRef.current = promptPath;
            setSftpPath(promptPath);
            void invoke("send_sftp_input", {
              sessionId: sftpSessionIdRef.current,
              input: `cd ${quoteSftpArg(promptPath)}\rpwd\rls -l\r`,
            }).catch((err) => setError(String(err)));
          }
          terminalRef.current?.write(payload.data);
        });
        if (disposed && unlisten) unlisten();
      } catch (error) {
        console.error("监听ssh-output事件失败:", error);
        terminalRef.current?.writeln(`\r\n[debug] 监听 ssh-output 失败: ${String(error)}\r\n`);
      }
    })();
    return () => {
      disposed = true;
      if (unlisten) unlisten();
    };
  }, []);

  useEffect(() => {
    // 只在Tauri环境中监听事件
    if (typeof window === "undefined" || !isTauri()) return;
    
    let unlisten: UnlistenFn | null = null;
    let disposed = false;
    void (async () => {
      try {
        console.log("[SFTP] registering sftp-output listener");
        unlisten = await listen<PtyOutputEvent>("sftp-output", (event) => {
          // 使用ref来比较，确保获取最新的sessionId
          if (event.payload.sessionId !== sftpSessionIdRef.current) {
            console.log("[SFTP] Session ID mismatch:", event.payload.sessionId, "!==", sftpSessionIdRef.current);
            return;
          }
          if (/password:/i.test(event.payload.data)) {
            const password = sharedPasswordRef.current;
            if (password && !sftpPasswordSentRef.current) {
              sftpPasswordSentRef.current = true;
              void invoke("send_sftp_input", {
                sessionId: event.payload.sessionId,
                input: `${password}\r`,
              }).catch((err) => setError(String(err)));
            }
          }
          if (/Connected to /i.test(event.payload.data)) {
            console.log("[SFTP] Connected to detected");
            sftpReadyRef.current = true;
            if (!sftpAutoListedRef.current) {
              sftpAutoListedRef.current = true;
              console.log("[SFTP] Sending ls -l command");
              // 使用setTimeout确保SFTP连接完全建立后再发送命令
              setTimeout(() => {
                void invoke("send_sftp_input", {
                  sessionId: event.payload.sessionId,
                  input: "ls -l\r",
                }).catch((err) => setError(String(err)));
              }, 100);
            }
          }
          const next = `${sftpRawRef.current}${event.payload.data}`.slice(-20000);
          sftpRawRef.current = next;
          const sanitizedNext = stripAnsiControls(next);
          const pwdMatches = Array.from(sanitizedNext.matchAll(/Remote working directory:\s*(.+)/gi));
          const latestPwd = pwdMatches.length > 0 ? pwdMatches[pwdMatches.length - 1][1] : null;
          if (latestPwd) {
            setSftpPath(latestPwd.trim());
          }
          const transferProgress = parseSftpTransferProgress(event.payload.data);
          if (transferProgress) {
            pushSftpProgress(transferProgress);
          }
          if (/sftp>\s*$/m.test(event.payload.data)) {
            scheduleSftpProgressAutoHide(900);
          }
          const entries = parseSftpEntries(next);
          console.log("[SFTP] Parsed entries:", entries);
          setSftpEntries(entries);
        });
        if (disposed && unlisten) unlisten();
      } catch (error) {
        console.error("监听sftp-output事件失败:", error);
      }
    })();
    return () => {
      disposed = true;
      if (unlisten) unlisten();
    };
  }, [sftpSessionId]);

  const canConnect = useMemo(
    () => address.trim().length > 0 && !sessionId && !isConnecting,
    [address, sessionId, isConnecting]
  );

  /** 建立 SSH + SFTP 连接并初始化终端状态。 / Connect SSH and SFTP, then initialize terminal state. */
  async function handleConnect() {
    try {
      connectStartedAtRef.current = Date.now();
      setIsConnecting(true);
      waitingForShellPromptRef.current = true;
      setError("");
      const effectivePassword = passwordRef.current || password;
      console.log("[CONNECT] handleConnect start", {
        address,
        hasPassword: Boolean(effectivePassword),
        passwordLength: effectivePassword.length,
        isTauri: isTauri(),
      });
      terminalRef.current?.writeln(`\r\n[debug] connecting: ${address}\r\n`);
      sftpRawRef.current = "";
      setSftpPath("/");
      sftpReadyRef.current = false;
      sftpTransferBusyRef.current = false;
      sftpAutoListedRef.current = false;
      sftpPasswordSentRef.current = false;
      sharedPasswordRef.current = effectivePassword;
      sshPasswordBufferRef.current = "";
      sshAwaitingPasswordRef.current = false;
      setSftpEntries([]);
      setStatus(t.statusConnecting);

      const sid = await invoke<string>("connect_ssh", { payload: { address, password: effectivePassword } });
      console.log("[CONNECT] ssh session created", sid);
      terminalRef.current?.writeln(`[debug] ssh sid=${sid}\r\n`);
      setSessionId(sid);
      sessionIdRef.current = sid;

      const sftpSid = await invoke<string>("connect_sftp", { payload: { address, password: effectivePassword } });
      console.log("[CONNECT] sftp session created", sftpSid);
      terminalRef.current?.writeln(`[debug] sftp sid=${sftpSid}\r\n`);
      setSftpSessionId(sftpSid);
      sftpSessionIdRef.current = sftpSid;
      console.log("[SFTP] Connected with session ID:", sftpSid);

      setStatus(tr("statusConnected", { sid }));
      terminalRef.current?.clear();
      terminalRef.current?.writeln(`[system] connecting to ${address}`);
      fitRef.current?.fit();
      void invoke("resize_ssh", {
        sessionId: sid,
        cols: terminalRef.current?.cols ?? 120,
        rows: terminalRef.current?.rows ?? 40,
      });
      // 连接成功后，终端切换到根目录，与SFTP保持同步
      await invoke("send_ssh_input", { 
        sessionId: sid, 
        input: `alias ls='ls --color=never' 2>/dev/null\r\ncd /\r\n${SSH_LIST_CMD}\r\n` 
      });
      terminalRef.current?.focus();
    } catch (err) {
      console.error("[CONNECT] failed", err);
      waitingForShellPromptRef.current = false;
      setSessionId(null);
      setSftpSessionId(null);
      sessionIdRef.current = null;
      setStatus(t.statusConnectFailed);
      setError(String(err));
      terminalRef.current?.writeln(`\r\n[debug] connect failed: ${String(err)}\r\n`);
      finishConnectingWithMinimumDelay();
    }
  }

  /** 断开 SSH 与 SFTP 连接，并清理界面状态。 / Disconnect SSH and SFTP, then reset UI state. */
  async function handleDisconnect() {
    const sid = sessionIdRef.current;
    const sftpSid = sftpSessionIdRef.current;
    try {
      if (sid) await invoke("disconnect_ssh", { sessionId: sid });
      if (sftpSid) await invoke("disconnect_sftp", { sessionId: sftpSid });
    } finally {
      setSessionId(null);
      sessionIdRef.current = null;
      setSftpSessionId(null);
      sftpSessionIdRef.current = null;
      setStatus(t.statusDisconnected);
      setSftpEntries([]);
      clearSftpProgressHideTimer();
      setSftpProgress(null);
      sftpRawRef.current = "";
      setSftpPath("/");
      sftpReadyRef.current = false;
      sftpAutoListedRef.current = false;
      sftpPasswordSentRef.current = false;
      sharedPasswordRef.current = "";
      sshPasswordBufferRef.current = "";
      sshAwaitingPasswordRef.current = false;
      waitingForShellPromptRef.current = false;
      setIsConnecting(false);
      sshOutputTailRef.current = "";
      lastSyncedPromptPathRef.current = "/";
      // 清空地址栏和密码
      setAddress("");
      setPassword("");
      // 清空终端并显示欢迎信息
      if (terminalRef.current) {
        terminalRef.current.clear();
        terminalRef.current.writeln(t.welcome);
        terminalRef.current.writeln(t.welcomeHint);
      }
    }
  }

  /** 刷新当前 SFTP 目录列表。 / Refresh current SFTP directory list. */
  async function refreshSftpDir() {
    if (!sftpSessionId) return;
    if (guardSftpBusy(appLanguage === "zh-CN" ? "刷新目录" : "refreshing directory")) return;
    try {
      if (!sftpReadyRef.current) {
        setError(appLanguage === "zh-CN" ? "SFTP 尚未认证完成，请先发送密码。" : "SFTP is not authenticated yet. Please send password first.");
        return;
      }
      await invoke("send_sftp_input", {
        sessionId: sftpSessionId,
        input: "pwd\rls -l\r",
      });
    } catch (err) {
      setError(String(err));
    }
  }

  /** 切换 SFTP 目录，并按需同步到 SSH 终端目录。 / Change SFTP directory and optionally sync SSH terminal path. */
  async function changeSftpDir(target: string) {
    if (!sftpSessionId || !sftpReadyRef.current) return;
    if (guardSftpBusy(appLanguage === "zh-CN" ? "切换目录" : "changing directory")) return;
    try {
      sftpRawRef.current = "";
      setSftpEntries([]);
      let newPath = target === ".." ? "" : target;
      if (newPath && !sftpPath.endsWith("/")) newPath = `/${newPath}`;
      newPath = `${sftpPath}${newPath}`;
      if (target === "..") {
        const parts = sftpPath.split("/").filter(Boolean);
        newPath = parts.slice(0, -1).join("/") || "/";
      }
      if (!newPath.startsWith("/")) newPath = `/${newPath}`;
      // 先在前端更新目标路径，避免与 SSH 同步回写产生“需要点两次”的竞态。
      setSftpPath(newPath);
      lastSyncedPromptPathRef.current = newPath;
      const shouldRetryForNewDir =
        target !== ".." &&
        (Date.now() - (recentlyCreatedSftpDirsRef.current.get(target) ?? 0) < 8000);
      await invoke("send_sftp_input", {
        sessionId: sftpSessionId,
        // 使用 \r 分隔，避免某些 Windows/OpenSSH 组合对 \n 多行解析不稳定。
        input: `cd ${quoteSftpArg(newPath)}\rpwd\rls -la\r`,
      });
      if (shouldRetryForNewDir) {
        recentlyCreatedSftpDirsRef.current.delete(target);
        window.setTimeout(() => {
          const sid = sftpSessionIdRef.current;
          if (!sid) return;
          void invoke("send_sftp_input", {
            sessionId: sid,
            input: `cd ${quoteSftpArg(newPath)}\rpwd\rls -la\r`,
          }).catch((err) => setError(String(err)));
        }, 180);
      }
      const sid = sessionIdRef.current;
      if (syncDirEnabled && sid) {
        let terminalCommand = "";
        if (target === "..") {
          terminalCommand = `cd ..\r\n${SSH_LIST_CMD}\r\n`;
        } else {
          terminalCommand = `cd ${quoteSftpArg(newPath)}\r\n${SSH_LIST_CMD}\r\n`;
        }
        await invoke("send_ssh_input", { sessionId: sid, input: terminalCommand });
      }
    } catch (err) {
      setError(String(err));
    }
  }

  /** 在当前 SFTP 目录创建新文件夹。 / Create a new folder in current SFTP directory. */
  async function createSftpDirByName(rawName: string) {
    const name = rawName.trim();
    if (!sftpSessionId) return;
    if (guardSftpBusy(appLanguage === "zh-CN" ? "创建文件夹" : "creating folder")) return;
    if (!sftpReadyRef.current) {
      setError(appLanguage === "zh-CN" ? "SFTP 尚未认证完成，请先发送密码。" : "SFTP is not authenticated yet. Please send password first.");
      return;
    }
    if (!name || name === "." || name === "..") {
      setError(t.invalidFolderName);
      return;
    }
    setError("");
    const q = quoteSftpArg(name);
    try {
      await invoke("send_sftp_input", {
        sessionId: sftpSessionId,
        input: `mkdir ${q}\rpwd\rls -la\r`,
      });
      recentlyCreatedSftpDirsRef.current.set(name, Date.now());
    } catch (err) {
      setError(String(err));
    }
  }

  /** 在当前目录创建空文件（通过 SSH 执行）。 / Create an empty file in current directory via SSH. */
  async function createSftpEmptyFileByName(rawName: string) {
    const name = rawName.trim();
    if (!sftpSessionId) return;
    if (guardSftpBusy(appLanguage === "zh-CN" ? "创建空文件" : "creating empty file")) return;
    if (!sftpReadyRef.current) {
      setError(appLanguage === "zh-CN" ? "SFTP 尚未认证完成，请先发送密码。" : "SFTP is not authenticated yet. Please send password first.");
      return;
    }
    if (!name || name === "." || name === ".." || name.includes("/")) {
      setError(t.invalidFileName);
      return;
    }
    setError("");
    const fullPath = sftpPath === "/" ? `/${name}` : `${sftpPath}/${name}`;
    const q = quoteSftpArg(fullPath);
    try {
      const sshSid = sessionIdRef.current;
      if (!sshSid) {
        setError(appLanguage === "zh-CN" ? "创建空文件需要 SSH 会话在线。" : "Creating empty file requires an active SSH session.");
        return;
      }
      // OpenSSH 的 sftp 子命令不支持 touch，这里通过 SSH 在远端创建 0 字节文件，再刷新 SFTP 列表。
      await invoke("send_ssh_input", {
        sessionId: sshSid,
        input: `: > ${q}\r\n`,
      });
      await invoke("send_sftp_input", {
        sessionId: sftpSessionId,
        input: "pwd\rls -la\r",
      });
    } catch (err) {
      setError(String(err));
    }
  }

  /** 打开 SFTP 路径栏右键菜单。 / Open context menu on SFTP path bar. */
  function handleSftpHeaderContextMenu(e: React.MouseEvent<HTMLDivElement>) {
    e.preventDefault();
    if (!sftpSessionId) return;
    setSftpHeaderMenu({ visible: true, x: e.clientX, y: e.clientY });
  }

  /** 处理会话项的右键菜单。 / Handle session item context menu. */
  function handleSessionContextMenu(e: React.MouseEvent, session: Session) {
    e.preventDefault();
    setSessionMenu({
      visible: true,
      x: e.clientX,
      y: e.clientY,
      session,
    });
  }

  /** 处理修改会话。 / Handle edit session. */
  function handleEditSession() {
    if (sessionMenu.session) {
      // 填充表单进行修改
      setAddress(`${sessionMenu.session.user}@${sessionMenu.session.host}:${sessionMenu.session.port}`);
      const selectedPassword = sessionMenu.session.password || "";
      passwordRef.current = selectedPassword;
      setPassword(selectedPassword);
      setNewSession({
        name: sessionMenu.session.name,
        host: sessionMenu.session.host,
        port: sessionMenu.session.port,
        user: sessionMenu.session.user,
        password: sessionMenu.session.password || "",
        group: sessionMenu.session.group
      });
      setShowNewSessionModal(true);
      // 关闭菜单
      setSessionMenu(prev => ({ ...prev, visible: false }));
    }
  }

  /** 处理复制会话。 / Handle copy session. */
  function handleCopySession() {
    if (sessionMenu.session) {
      const newSession: Session = {
        ...sessionMenu.session,
        id: `session-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
        name: `${sessionMenu.session.name} (副本)`,
      };
      const updatedSessions = [...sessions, newSession];
      setSessions(updatedSessions);
      void saveSessionsToFile(updatedSessions);
      // 关闭菜单
      setSessionMenu(prev => ({ ...prev, visible: false }));
    }
  }

  /** 处理删除会话。 / Handle delete session. */
  function handleDeleteSession() {
    if (sessionMenu.session) {
      const updatedSessions = sessions.filter(s => s.id !== sessionMenu.session?.id);
      setSessions(updatedSessions);
      void saveSessionsToFile(updatedSessions);
      // 关闭菜单
      setSessionMenu(prev => ({ ...prev, visible: false }));
    }
  }

  /** 处理连接会话。 / Handle connect session. */
  function handleConnectSession() {
    if (sessionMenu.session) {
      // 填充连接信息
      setAddress(`${sessionMenu.session.user}@${sessionMenu.session.host}:${sessionMenu.session.port}`);
      const selectedPassword = sessionMenu.session.password || "";
      passwordRef.current = selectedPassword;
      setPassword(selectedPassword);
      // 关闭菜单
      setSessionMenu(prev => ({ ...prev, visible: false }));
      // 执行连接
      void handleConnect();
    }
  }

  /** 处理会话列表空白区域的右键菜单。 / Handle session list context menu. */
  function handleSessionListContextMenu(e: React.MouseEvent) {
    e.preventDefault();
    setSessionListMenu({
      visible: true,
      x: e.clientX,
      y: e.clientY,
    });
  }

  /** 处理创建群组。 / Handle create group. */
  function handleCreateGroup() {
    const groupName = window.prompt("请输入群组名称：", "新群组");
    if (groupName && groupName.trim()) {
      // 创建一个临时会话来初始化新群组
      const tempSession: Session = {
        id: `temp-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
        name: "临时会话",
        host: "localhost",
        port: 22,
        user: "root",
        password: "",
        group: groupName.trim()
      };
      
      // 添加临时会话
      const tempSessions = [...sessions, tempSession];
      setSessions(tempSessions);
      
      // 立即删除临时会话，只保留群组结构
      setTimeout(() => {
        const updatedSessions = tempSessions.filter(s => s.id !== tempSession.id);
        setSessions(updatedSessions);
        void saveSessionsToFile(updatedSessions);
      }, 100);
      
      // 关闭菜单
      setSessionListMenu(prev => ({ ...prev, visible: false }));
    }
  }

  /** 处理会话项的拖拽开始。 / Handle session item drag start. */
  function handleSessionDragStart(e: React.DragEvent, session: Session) {
    e.dataTransfer.setData('application/json', JSON.stringify(session));
  }

  /** 处理群组的拖放。 / Handle group drop. */
  function handleGroupDrop(e: React.DragEvent, group: string) {
    e.preventDefault();
    const sessionJson = e.dataTransfer.getData('application/json');
    if (sessionJson) {
      const session = JSON.parse(sessionJson) as Session;
      if (session.group !== group) {
        const updatedSessions = sessions.map(s => 
          s.id === session.id ? { ...s, group } : s
        );
        setSessions(updatedSessions);
        void saveSessionsToFile(updatedSessions);
      }
    }
  }

  /** 处理群组的拖拽悬停。 / Handle group drag over. */
  function handleGroupDragOver(e: React.DragEvent) {
    e.preventDefault();
  }

  /** 通过右键菜单触发“创建文件夹”。 / Trigger create-folder action from context menu. */
  function handleSftpContextCreateDir() {
    setSftpHeaderMenu((prev) => ({ ...prev, visible: false }));
    const name = window.prompt(t.inputFolderNamePrompt, "");
    if (name == null) return;
    void createSftpDirByName(name);
  }

  /** 通过右键菜单触发“创建空文件”。 / Trigger create-empty-file action from context menu. */
  function handleSftpContextCreateFile() {
    setSftpHeaderMenu((prev) => ({ ...prev, visible: false }));
    const name = window.prompt(t.inputFileNamePrompt, "");
    if (name == null) return;
    void createSftpEmptyFileByName(name);
  }

  /** 点击图标按钮创建文件夹。 / Create folder from icon action button. */
  function handleSftpCreateDirIconClick() {
    const name = window.prompt(t.inputFolderNamePrompt, "");
    if (name == null) return;
    void createSftpDirByName(name);
  }

  /** 点击图标按钮创建空文件。 / Create empty file from icon action button. */
  function handleSftpCreateFileIconClick() {
    const name = window.prompt(t.inputFileNamePrompt, "");
    if (name == null) return;
    void createSftpEmptyFileByName(name);
  }

  /** 删除 SFTP 条目（文件或目录）并刷新列表。 / Delete SFTP entry (file/folder) and refresh list. */
  async function deleteSftpEntry(entry: SftpEntry) {
    if (!sftpSessionId) return;
    if (guardSftpBusy(appLanguage === "zh-CN" ? "删除条目" : "deleting item")) return;
    if (!sftpReadyRef.current) {
      setError(appLanguage === "zh-CN" ? "SFTP 尚未认证完成，请先发送密码。" : "SFTP is not authenticated yet. Please send password first.");
      return;
    }
    const kind = entry.isDir ? t.dir : t.file;
    if (!window.confirm(tr("deleteConfirm", { kind, name: entry.name }))) return;
    setError("");
    const fullPath = sftpPath === "/" ? `/${entry.name}` : `${sftpPath}/${entry.name}`;
    const q = quoteSftpArg(fullPath);
    try {
      const sshSid = sessionIdRef.current;
      if (!sshSid) {
        setError(appLanguage === "zh-CN" ? "删除需要 SSH 会话在线。" : "Deleting requires an active SSH session.");
        return;
      }
      const cmd = entry.isDir ? `rm -rf ${q}` : `rm -f ${q}`;
      await invoke("send_ssh_input", {
        sessionId: sshSid,
        input: `${cmd}\r\n`,
      });
      // SSH 删除是异步执行，先清空当前列表，再稍后刷新，避免界面短暂显示旧目录。
      sftpRawRef.current = "";
      setSftpEntries([]);
      window.setTimeout(() => {
        void invoke("send_sftp_input", {
          sessionId: sftpSessionId,
          input: "pwd\rls -la\r",
        }).catch((refreshErr) => setError(String(refreshErr)));
      }, 180);
    } catch (err) {
      setError(String(err));
    }
  }

  /** 允许文件拖拽进入 SFTP 列表区域。 / Allow file drag over SFTP list area. */
  function handleDragOver(e: React.DragEvent) {
    e.preventDefault();
    const uriList = e.dataTransfer.getData("text/uri-list");
    const hasLocalFiles = e.dataTransfer.files.length > 0 || Boolean(uriList.trim());
    e.dataTransfer.dropEffect = hasLocalFiles ? "copy" : "move";
  }

  /** 提取拖拽中的本地文件路径（兼容 Tauri / 浏览器不同数据格式）。 / Extract dropped local paths from multiple drag payload formats. */
  function getDroppedLocalPaths(e: React.DragEvent): string[] {
    const paths: string[] = [];
    const files = Array.from(e.dataTransfer.files);
    for (const file of files) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const localPath = (file as any).path as string | undefined;
      if (localPath) paths.push(localPath);
    }
    const uriList = e.dataTransfer.getData("text/uri-list");
    if (uriList) {
      const uris = uriList.split(/\r?\n/).map((v) => v.trim()).filter((v) => v && !v.startsWith("#"));
      for (const uri of uris) {
        if (!uri.toLowerCase().startsWith("file://")) continue;
        try {
          const url = new URL(uri);
          let p = decodeURIComponent(url.pathname);
          if (/^\/[A-Za-z]:\//.test(p)) p = p.slice(1);
          p = p.replace(/\//g, "\\");
          if (p) paths.push(p);
        } catch {
          // ignore malformed uri
        }
      }
    }
    return Array.from(new Set(paths));
  }

  /** 上传本地路径列表到当前 SFTP 目录并显示进度。 / Upload local file paths to current SFTP directory with progress. */
  async function uploadLocalPaths(localPaths: string[]) {
    if (!sftpSessionIdRef.current || !sftpReadyRef.current) return;
    sftpTransferBusyRef.current = true;
    for (const localPath of localPaths) {
      const fileName = localPath.split("\\").pop() || localPath;
      pushSftpProgress({
        file: fileName,
        percent: 0,
        speed: "--",
      });
      const cmd = `put "${localPath}" "${sftpPathRef.current}"\r`;
      try {
        await invoke("send_sftp_input", { sessionId: sftpSessionIdRef.current, input: cmd });
        // 每次上传后刷新目录，确保新文件立即出现在列表中。
        await invoke("send_sftp_input", {
          sessionId: sftpSessionIdRef.current,
          input: "pwd\rls -la\r",
        });
      } catch (err) {
        setError(String(err));
      }
    }
    sftpTransferBusyRef.current = false;
  }

  /** 调用原生拖拽命令（按下即进入系统拖拽）。 / Invoke native drag directly on mouse down. */
  async function invokeNativeDragOnce(
    payload: {
      sftpSessionId: string;
      remotePath: string;
      displayName: string;
      isDir: boolean;
    },
    traceId: string
  ): Promise<{ effect: number }> {
    pushDebugLog(`[M4][${traceId}] invoke begin item=${payload.displayName} isDir=${payload.isDir}`);
    pushDebugLog(`[M4][${traceId}] use pick-target+background-download`);
    const savedPath = await invoke<string>("native_pick_drop_target_and_download", {
      payload: {
        sftpSessionId: payload.sftpSessionId,
        remotePath: payload.remotePath,
        displayName: payload.displayName,
        isDir: payload.isDir,
      },
    });
    pushDebugLog(`[M4][${traceId}] savedPath=${savedPath}`);
    return { effect: 1 };
  }

  /** 鼠标按下后触发原生拖拽（不在按下阶段做缓存下载）。 / Start native drag on mouse down without pre-cache download. */
  function handleFileMouseDown(e: React.MouseEvent, entry: SftpEntry) {
    if (!isTauri() || !sftpSessionIdRef.current || !sftpReadyRef.current) return;
    if (e.button !== 0) return;
    if (nativeDragRunningRef.current) return;
    const remotePath = sftpPathRef.current === "/" ? `/${entry.name}` : `${sftpPathRef.current}/${entry.name}`;
    dragTraceSeqRef.current += 1;
    const traceId = `${Date.now()}-${dragTraceSeqRef.current}`;
    e.preventDefault();
    pushDebugLog(`[M4][${traceId}] mousedown item=${entry.name} remote=${remotePath}`);
    setNativeDragStage("preparing", entry.name, remotePath, `dragTrace=${traceId}`);
    nativeDragRunningRef.current = true;
    sftpTransferBusyRef.current = true;
    setGlobalDragCursor(true);
    pushSftpProgress({
      file: entry.name,
      percent: 1,
      speed: "正在系统拖拽，请在目标目录松手...",
    });
    setDraggingExportName(entry.name);
    void invokeNativeDragOnce({
      sftpSessionId: sftpSessionIdRef.current,
      remotePath,
      displayName: entry.name,
      isDir: entry.isDir,
    }, traceId)
      .then((res) => {
        pushDebugLog(`[M4][${traceId}] drag success item=${entry.name} effect=${res.effect}`);
        setNativeDragStatus({ stage: "idle", item: "", targetPath: "", message: "" });
        if ((res.effect ?? 0) !== 0) {
          pushSftpProgress({
            file: entry.name,
            percent: 100,
            speed: "完成",
          });
          void invoke("send_sftp_input", {
            sessionId: sftpSessionIdRef.current,
            input: "pwd\rls -la\r",
          });
        }
      })
      .catch((err) => {
        pushDebugLog(`[M4][${traceId}] drag failed err=${String(err)}`);
        setNativeDragStage("failed", entry.name, "", "原生拖拽失败，请查看错误信息");
        setError(String(err));
        scheduleSftpProgressAutoHide(300);
      })
      .finally(() => {
        pushDebugLog(`[M4][${traceId}] drag finalize`);
        nativeDragRunningRef.current = false;
        sftpTransferBusyRef.current = false;
        setGlobalDragCursor(false);
        setDraggingExportName("");
      });
  }

  /** 处理拖拽上传：本地文件放入列表后上传。 / Handle drag upload from local files. */
  async function handleDrop(e: React.DragEvent) {
    e.preventDefault();
    if (!sftpSessionId || !sftpReadyRef.current) return;
    const localPaths = getDroppedLocalPaths(e);
    if (localPaths.length > 0) {
      await uploadLocalPaths(localPaths);
    }
  }

  useEffect(() => {
    if (typeof window === "undefined" || !isTauri()) return;
    const onMouseUp = () => {
      setGlobalDragCursor(false);
      setDraggingExportName("");
    };
    const onBlur = () => {
      setGlobalDragCursor(false);
      setDraggingExportName("");
    };
    window.addEventListener("mouseup", onMouseUp);
    window.addEventListener("blur", onBlur);
    let unlistenDrop: (() => void) | null = null;
    let disposed = false;
    void (async () => {
      try {
        unlistenDrop = await getCurrentWindow().onDragDropEvent((event) => {
          if (disposed) return;
          if (event.payload.type !== "drop") return;
          const paths = (event.payload.paths || []).filter((p) => !!p);
          if (paths.length === 0) return;
          void uploadLocalPaths(paths);
        });
      } catch {
        // ignore unsupported platforms/webviews
      }
    })();
    return () => {
      disposed = true;
      setGlobalDragCursor(false);
      window.removeEventListener("mouseup", onMouseUp);
      window.removeEventListener("blur", onBlur);
      if (unlistenDrop) unlistenDrop();
    };
  }, []);

  /** 打开“新建会话”弹窗。 / Open the new-session modal. */
  function handleNewSession() {
    setShowNewSessionModal(true);
  }

  /** 新建会话：写入列表并自动保存到 mmshell_config.json（含用户名、主机、端口、密码等）。 / Create or update a session and persist it to config. */
  async function handleSaveSession() {
    const name = newSession.name.trim();
    const host = newSession.host.trim();
    if (!name || !host) {
      setError(appLanguage === "zh-CN" ? "请填写会话名称和主机地址" : "Please fill session name and host address.");
      return;
    }
    setError("");
    const duplicate = sessions.find(
      (s) => s.host === host && s.port === newSession.port && s.user === newSession.user
    );
    const entry: Session = {
      ...newSession,
      name,
      host,
      id: duplicate?.id ?? Date.now().toString(),
    };
    const updatedSessions = duplicate
      ? sessions.map((s) => (s.id === duplicate.id ? entry : s))
      : [...sessions, entry];
    setSessions(updatedSessions);
    const ok = await saveSessionsToFile(updatedSessions);
    if (ok) {
      setShowNewSessionModal(false);
      setNewSession({
        name: "",
        host: "",
        port: 22,
        user: "root",
        password: "",
        group: t.groupCommon,
      });
      setStatus(t.statusSessionSaved);
    } else {
      setError(appLanguage === "zh-CN" ? "保存失败：请在 Tauri 桌面版中运行，并确认可写入 D:\\MMShell0414" : "Save failed: run in Tauri desktop and ensure D:\\MMShell0414 is writable.");
    }
  }

  /** 将工具栏当前 user@host:port 与密码保存为会话。 / Quickly save current toolbar connection as a session. */
  async function handleSaveCurrentConnection() {
    const parsed = parseConnectionAddress(address);
    if (!parsed) {
      setError(appLanguage === "zh-CN" ? "地址格式应为 user@host 或 user@host:port" : "Address format should be user@host or user@host:port.");
      return;
    }
    setError("");
    const defaultName = `${parsed.user}@${parsed.host}`;
    const duplicate = sessions.find(
      (s) => s.host === parsed.host && s.port === parsed.port && s.user === parsed.user
    );
    const entry: Session = {
      id: duplicate?.id ?? Date.now().toString(),
      name: duplicate?.name ?? defaultName,
      host: parsed.host,
      port: parsed.port,
      user: parsed.user,
      password,
      group: duplicate?.group ?? t.groupCommon,
    };
    const updatedSessions = duplicate
      ? sessions.map((s) => (s.id === duplicate.id ? entry : s))
      : [...sessions, entry];
    setSessions(updatedSessions);
    const ok = await saveSessionsToFile(updatedSessions);
    if (ok) {
      setStatus(t.statusCurrentSaved);
    } else {
      setError(t.saveFailed);
    }
  }

  // 保存会话到本地文件（Tauri + plugin-fs）；主题配色变更时也会写入 settings.terminalTheme
  async function saveSessionsToFile(
    sessionsToSave: Session[],
    languageOverride?: AppLanguage,
    themeSnapshot?: { presetId: ThemePresetId; theme: TerminalTheme }
  ): Promise<boolean> {
    try {
      if (typeof window === "undefined" || !isTauri()) return false;
      const preset = themeSnapshot?.presetId ?? themePresetId;
      const theme = themeSnapshot?.theme ?? customTheme;
      const config = {
        version: 1,
        sessions: sessionsToSave,
        settings: {
          language: languageOverride ?? appLanguage,
          themePresetId: preset,
          terminalTheme: theme,
        },
        lastUpdated: new Date().toISOString(),
      };
      await writeTextFile(SESSIONS_CONFIG_PATH, JSON.stringify(config, null, 2));
      console.log("[Config] 会话配置已保存到", SESSIONS_CONFIG_PATH);
      return true;
    } catch (error) {
      console.error("保存会话失败:", error);
      return false;
    }
  }

  // 从本地文件加载会话 - 从项目根目录加载
  async function loadSessionsFromFile() {
    try {
      if (!isTauri()) return;
      if (!(await exists(SESSIONS_CONFIG_PATH))) return;
      const contents = await readTextFile(SESSIONS_CONFIG_PATH);
      const config = JSON.parse(contents) as {
        sessions?: Session[];
        settings?: {
          language?: AppLanguage;
          themePresetId?: string;
          terminalTheme?: Partial<TerminalTheme>;
        };
      };
      if (config.sessions && Array.isArray(config.sessions)) {
        setSessions(config.sessions);
        console.log("[Config] 会话配置已从", SESSIONS_CONFIG_PATH, "加载");
      }
      if (config.settings?.language && (config.settings.language === "zh-CN" || config.settings.language === "en-US")) {
        setAppLanguage(config.settings.language);
      }
      if (config.settings?.themePresetId && isThemePresetId(config.settings.themePresetId)) {
        setThemePresetId(config.settings.themePresetId);
      }
      if (config.settings?.terminalTheme && typeof config.settings.terminalTheme === "object") {
        setCustomTheme(mergeLoadedTerminalTheme(config.settings.terminalTheme));
      }
    } catch (error) {
      console.error("加载会话失败:", error);
    }
  }

  /** 处理会话列表点击，回填地址与密码。 / Handle session form change and update draft fields. */
  function handleSessionChange(e: React.ChangeEvent<HTMLInputElement>) {
    const { name, value } = e.target;
    setNewSession(prev => ({
      ...prev,
      [name]: name === "port" ? parseInt(value) || 22 : value
    }));
  }

  /** 切换终端主题预设并立即持久化。 / Switch terminal theme preset and persist immediately. */
  function handleThemePresetChange(e: React.ChangeEvent<HTMLSelectElement>) {
    const nextId = e.target.value as ThemePresetId;
    setThemePresetId(nextId);
    if (nextId !== "custom") {
      const newTheme = TERMINAL_THEME_PRESETS[nextId].theme;
      setCustomTheme(newTheme);
      void saveSessionsToFile(sessions, undefined, { presetId: nextId, theme: newTheme });
    } else {
      void saveSessionsToFile(sessions, undefined, { presetId: "custom", theme: customTheme });
    }
  }

  /** 修改自定义主题单个颜色并保存。 / Update one custom theme color and save. */
  function handleCustomThemeColorChange(key: keyof TerminalTheme, color: string) {
    setThemePresetId("custom");
    setCustomTheme((prev) => {
      const next = { ...prev, [key]: color };
      void saveSessionsToFile(sessions, undefined, { presetId: "custom", theme: next });
      return next;
    });
  }

  /** 将自定义主题重置为 Shell Dark 配色。 / Reset custom theme to Shell Dark palette. */
  function handleResetCustomTheme() {
    const fallback = TERMINAL_THEME_PRESETS["shell-dark"].theme;
    setThemePresetId("custom");
    setCustomTheme(fallback);
    void saveSessionsToFile(sessions, undefined, { presetId: "custom", theme: fallback });
  }

  /** 切换应用语言并写入本地配置。 / Change app language and persist to local config. */
  async function handleLanguageChange(e: React.ChangeEvent<HTMLSelectElement>) {
    const next = e.target.value as AppLanguage;
    setAppLanguage(next);
    const ok = await saveSessionsToFile(sessions, next);
    if (!ok) {
      setError(I18N[next].saveFailed);
    }
  }

  const groupedSessions = sessions.reduce((groups, session) => {
    if (!groups[session.group]) {
      groups[session.group] = [];
    }
    groups[session.group].push(session);
    return groups;
  }, {} as Record<string, Session[]>);

  return (
    <>
      <main className="app shell-style">
        {/* 菜单栏 */}
        <div className="shell-menu-bar">
          <div className="menu-left">
            <div className="menu-item">
              {t.menuFile}
              <div className="submenu">
                <div className="submenu-item">{t.newSession}</div>
                <div className="submenu-item">{t.openSession}</div>
                <div className="submenu-divider"></div>
                <div className="submenu-item">{t.exit}</div>
              </div>
            </div>
            <div className="menu-item">
              {t.menuEdit}
              <div className="submenu">
                <div className="submenu-item">{t.copy}</div>
                <div className="submenu-item">{t.paste}</div>
              </div>
            </div>
            <div className="menu-item">
              {t.menuView}
              <div className="submenu">
                <div className="submenu-item">{t.fullscreen}</div>
                <div className="submenu-item">{t.splitWindow}</div>
              </div>
            </div>
            <div className="menu-item">
              {t.menuTools}
              <div className="submenu">
                <div className="submenu-item" onClick={() => setShowSettingsModal(true)}>{t.options}</div>
                <div className="submenu-item">{t.sshKeyManager}</div>
              </div>
            </div>
            <div className="menu-item">
              {t.menuHelp}
              <div className="submenu">
                <div className="submenu-item">{t.about}</div>
                <div className="submenu-item">{t.helpDocs}</div>
              </div>
            </div>
          </div>
          <div className="menu-right">
            <div className="window-controls">
              <button className="window-control"><Minus size={14} /></button>
              <button className="window-control"><Maximize2 size={14} /></button>
              <button className="window-control close"><X size={14} /></button>
            </div>
          </div>
        </div>

        {/* 工具栏 */}
        <div className="shell-toolbar">
          <div className="toolbar-left">
            <button className="toolbar-button" onClick={() => void handleConnect()} disabled={!canConnect}>{t.connect}</button>
            <button className="toolbar-button" onClick={() => void handleDisconnect()} disabled={!sessionId && !sftpSessionId}>{t.disconnect}</button>
            <div className="toolbar-separator"></div>
            <button className="toolbar-button" onClick={() => void refreshSftpDir()} disabled={!sftpSessionId}>{t.refresh}</button>
            <div className="toolbar-separator"></div>
            <button className="toolbar-button" onClick={handleNewSession}>{t.newSession}</button>
            <button className="toolbar-button" onClick={() => void handleSaveCurrentConnection()}>{t.saveSession}</button>
            <button className="toolbar-button" onClick={() => setShowThemeModal(true)}>
              <Palette size={14} />
              {t.theme}
            </button>
            <button className="toolbar-button" onClick={() => setShowSettingsModal(true)}>{t.settings}</button>
          </div>
          <div className="toolbar-right">
            <div className="connection-form">
              <input
                value={address}
                onChange={(e) => setAddress(e.currentTarget.value)}
                placeholder="user@host:port"
                className="address"
                disabled={Boolean(sessionId)}
              />
            </div>
          </div>
        </div>

        {/* 主内容区域 */}
        <div className="shell-main">
          {/* 左侧会话管理器 */}
          <div className={`shell-session-manager ${sessionPanelCollapsed ? 'collapsed' : ''} ${directoryPanelCollapsed ? 'directory-collapsed' : ''}`}>
            <div className="session-manager-container">
              {/* 会话面板 */}
              <div className={sessionPanelCollapsed ? 'session-panel-collapsed' : 'session-panel'}>
                <div 
                  className={`session-manager-header ${sessionPanelCollapsed ? 'collapsed' : ''}`}
                  onClick={() => setSessionPanelCollapsed(!sessionPanelCollapsed)}
                >
                  {!sessionPanelCollapsed ? (
                    <>
                      {t.session}
                      <ChevronLeft className="collapse-icon" />
                    </>
                  ) : (
                    <>
                      <Star className="star-icon" />
                      <span className="icon-label">{t.session}</span>
                    </>
                  )}
                </div>
                {!sessionPanelCollapsed && (
                  <div className="session-list" onContextMenu={handleSessionListContextMenu}>
                    {Object.entries(groupedSessions).map(([group, sessions]) => (
                      <div 
                        key={group} 
                        className="session-group"
                        onDrop={(e) => handleGroupDrop(e, group)}
                        onDragOver={handleGroupDragOver}
                      >
                        <div className="session-group-header">{displayGroupName(group)}</div>
                        {sessions.map((session) => (
                          <div 
                            key={session.id} 
                            className="session-item"
                            draggable
                            onDragStart={(e) => handleSessionDragStart(e, session)}
                            onClick={() => {
                              setAddress(`${session.user}@${session.host}:${session.port}`);
                              const selectedPassword = session.password || "";
                              passwordRef.current = selectedPassword;
                              setPassword(selectedPassword);
                            }}
                            onContextMenu={(e) => handleSessionContextMenu(e, session)}
                          >
                            {session.name}
                          </div>
                        ))}
                      </div>
                    ))}
                    {sessionMenu.visible && sessionMenu.session && (
                      <div
                        className="session-context-menu"
                        style={{ top: sessionMenu.y, left: sessionMenu.x }}
                        onClick={(e) => e.stopPropagation()}
                      >
                        <button type="button" className="session-context-item" onClick={handleConnectSession}>
                          连接
                        </button>
                        <button type="button" className="session-context-item" onClick={handleEditSession}>
                          修改
                        </button>
                        <button type="button" className="session-context-item" onClick={handleCopySession}>
                          复制
                        </button>
                        <button type="button" className="session-context-item" onClick={handleDeleteSession}>
                          删除
                        </button>
                      </div>
                    )}
                    {sessionListMenu.visible && (
                      <div
                        className="session-context-menu"
                        style={{ top: sessionListMenu.y, left: sessionListMenu.x }}
                        onClick={(e) => e.stopPropagation()}
                      >
                        <button type="button" className="session-context-item" onClick={handleCreateGroup}>
                          创建群组
                        </button>
                      </div>
                    )}
                  </div>
                )}
              </div>

              {/* 目录面板 */}
              <div className={directoryPanelCollapsed ? 'directory-panel-collapsed' : 'directory-panel'}>
                <div 
                  className={`directory-manager-header ${directoryPanelCollapsed ? 'collapsed' : ''}`}
                  onClick={() => setDirectoryPanelCollapsed(!directoryPanelCollapsed)}
                >
                  {!directoryPanelCollapsed ? (
                    <>
                      {t.sftpDirectory}
                      <ChevronLeft className="collapse-icon" />
                    </>
                  ) : (
                    <>
                      <FolderTree className="directory-icon" />
                      <span className="icon-label">{t.directory}</span>
                    </>
                  )}
                </div>
                {!directoryPanelCollapsed && (
                  <div className="directory-content">
                    <p className="hint">{t.sftpAuthHint}</p>
                    <div className="sftp-path-bar" onContextMenu={handleSftpHeaderContextMenu}>
                      {t.currentPath}：{sftpPath}
                    </div>
                    <div className="sftp-actions">
                      <button
                        type="button"
                        className="sftp-action-btn"
                        title={t.createDir}
                        aria-label={t.createDir}
                        onClick={handleSftpCreateDirIconClick}
                        disabled={!sftpSessionId}
                      >
                        <FolderPlus size={16} />
                      </button>
                      <button
                        type="button"
                        className="sftp-action-btn"
                        title={t.createEmptyFile}
                        aria-label={t.createEmptyFile}
                        onClick={handleSftpCreateFileIconClick}
                        disabled={!sftpSessionId}
                      >
                        <FilePlus2 size={16} />
                      </button>
                    </div>
                    {sftpHeaderMenu.visible && (
                      <div
                        className="sftp-context-menu"
                        style={{ top: sftpHeaderMenu.y, left: sftpHeaderMenu.x }}
                        onClick={(e) => e.stopPropagation()}
                      >
                        <button type="button" className="sftp-context-item" onClick={handleSftpContextCreateDir}>
                          {t.sftpContextCreateDir}
                        </button>
                        <button type="button" className="sftp-context-item" onClick={handleSftpContextCreateFile}>
                          {t.sftpContextCreateFile}
                        </button>
                      </div>
                    )}
                    <div className="sftp-sync-option">
                      <label>
                        <input
                          type="checkbox"
                          checked={syncDirEnabled}
                          onChange={(e) => setSyncDirEnabled(e.target.checked)}
                          disabled={!sftpSessionId}
                        />
                        {t.syncDirTerminal}
                      </label>
                    </div>
                    <ul className="sftp-list" onDragOver={handleDragOver} onDrop={handleDrop}>
                      {sftpEntries.length === 0 ? <li className="hint">{t.autoLoadHint}</li> : null}
                      {sftpSessionId && (
                        <li className="sftp-row">
                          <div className="sftp-item-main">
                            <button className="sftp-item dir" onClick={() => void changeSftpDir("..")}>
                              <Folder size={15} color="#1d4f91" />
                              <span className="sftp-name">..</span>
                            </button>
                          </div>
                        </li>
                      )}
                      {sftpEntries.map((entry) => (
                        <li key={entry.name} className="sftp-row">
                          <div className="sftp-item-main">
                            {entry.isDir ? (
                              <button className="sftp-item dir" onClick={() => void changeSftpDir(entry.name)}>
                                <Folder size={15} color={getSftpIconColorByName(entry.name, true)} />
                                <span className="sftp-name">{entry.name}</span>
                              </button>
                            ) : (
                              (() => {
                                const Icon = getSftpIconByName(entry.name);
                                return (
                                  <span
                                    className={`sftp-item file${draggingExportName === entry.name ? " dragging" : ""}`}
                                    draggable={false}
                                    onMouseDown={(e) => handleFileMouseDown(e, entry)}
                                  >
                                    <Icon size={15} color={getSftpIconColorByName(entry.name, false)} />
                                    <span className="sftp-name">{entry.name}</span>
                                  </span>
                                );
                              })()
                            )}
                          </div>
                          <button
                            type="button"
                            className="sftp-del"
                            title={t.delete}
                            aria-label={`${t.delete} ${entry.name}`}
                            onClick={() => void deleteSftpEntry(entry)}
                          >
                            <Trash2 size={14} />
                          </button>
                        </li>
                      ))}
                    </ul>
                    {sftpProgress && (
                      <div className="sftp-progress">
                        <div className="sftp-progress-info">
                          <span className="sftp-progress-file">{sftpProgress.file}</span>
                          <div style={{ display: 'flex', alignItems: 'center' }}>
                            <span className="sftp-progress-speed">{sftpProgress.speed}</span>
                          </div>
                        </div>
                        <div className="sftp-progress-bar-container">
                          <div className="sftp-progress-bar">
                            <div className="sftp-progress-fill" style={{ width: `${sftpProgress.percent}%` }} />
                          </div>
                          <span className="sftp-progress-percent">{sftpProgress.percent}%</span>
                        </div>
                      </div>
                    )}
                    {nativeDragStatus.stage === "failed" && (
                      <div className="hint">
                        原生拖拽状态：{nativeDragStatus.stage}｜对象：{nativeDragStatus.item}
                        {nativeDragStatus.message ? `｜${nativeDragStatus.message}` : ""}
                      </div>
                    )}
                  </div>
                )}
              </div>
            </div>
          </div>

          {/* 右侧内容区域 */}
          <div className="shell-content">
            {/* 标签页 */}
            <div className="shell-tabs">
              <div className="tab active">
                {tabTitle}
                <button className="tab-close" onClick={() => void handleDisconnect()}>×</button>
              </div>
            </div>

            {/* 终端容器 */}
            <div className="shell-terminal-container">
              {error && <p className="error">{error}</p>}
              <div className="terminal" ref={termHostRef} />
            </div>

            {/* 状态栏 */}
            <div className="shell-status-bar">
              <div className="status-left">
                <span className="status-item">{t.status}：{status}</span>
                <span className="status-item">{t.shortcut}：Ctrl+C / Ctrl+V</span>
              </div>
              <div className="status-right">
                <span className="status-item">{t.encoding}：UTF-8</span>
                <span className="status-item">{t.terminalType}：xterm-256color</span>
              </div>
            </div>
          </div>
        </div>
      </main>

      {/* 全局设置 */}
      {showSettingsModal && (
        <div className="modal-overlay">
          <div className="modal-content">
            <div className="modal-header">
              <h3>{t.settings}</h3>
              <button type="button" className="modal-close" onClick={() => setShowSettingsModal(false)}>×</button>
            </div>
            <div className="modal-body">
              <div className="form-group">
                <label>{t.language}</label>
                <select className="theme-select" value={appLanguage} onChange={(e) => void handleLanguageChange(e)}>
                  <option value="zh-CN">{t.chineseSimplified}</option>
                  <option value="en-US">{t.english}</option>
                </select>
                <p className="hint">{t.languageTip}</p>
              </div>
            </div>
            <div className="modal-footer">
              <button type="button" className="modal-button save" onClick={() => setShowSettingsModal(false)}>
                {t.close}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 终端主题设置 */}
      {showThemeModal && (
        <div className="modal-overlay">
          <div className="modal-content theme-modal">
            <div className="modal-header">
              <h3>{t.themeTitle}</h3>
              <button type="button" className="modal-close" onClick={() => setShowThemeModal(false)}>×</button>
            </div>
            <div className="modal-body">
              <div className="form-group">
                <label>{t.themePreset}</label>
                <select className="theme-select" value={themePresetId} onChange={handleThemePresetChange}>
                  <option value="shell-dark">{TERMINAL_THEME_PRESETS["shell-dark"].label}</option>
                  <option value="linux-classic">{TERMINAL_THEME_PRESETS["linux-classic"].label}</option>
                  <option value="one-dark">{TERMINAL_THEME_PRESETS["one-dark"].label}</option>
                  <option value="solarized-dark">{TERMINAL_THEME_PRESETS["solarized-dark"].label}</option>
                  <option value="monokai">{TERMINAL_THEME_PRESETS.monokai.label}</option>
                  <option value="custom">{t.custom}</option>
                </select>
              </div>

              <div className="theme-grid">
                {CUSTOM_THEME_FIELDS.map((item) => (
                  <label key={item.key} className="theme-color-item">
                    <span className="theme-color-label">
                      <strong>{item.label[appLanguage]}</strong>
                      <small>{item.description[appLanguage]}</small>
                    </span>
                    <div className="theme-color-control">
                      <input
                        type="color"
                        value={customTheme[item.key]}
                        onChange={(e) => handleCustomThemeColorChange(item.key, e.target.value)}
                      />
                      <code>{customTheme[item.key]}</code>
                    </div>
                  </label>
                ))}
              </div>
            </div>
            <div className="modal-footer">
              <button type="button" className="modal-button" onClick={handleResetCustomTheme}>{t.resetShellDark}</button>
              <button type="button" className="modal-button save" onClick={() => setShowThemeModal(false)}>{t.done}</button>
            </div>
          </div>
        </div>
      )}

      {/* 新建会话模态窗口 */}
      {showNewSessionModal && (
        <div className="modal-overlay">
          <div className="modal-content">
            <div className="modal-header">
              <h3>{t.newSessionTitle}</h3>
              <button type="button" className="modal-close" onClick={() => setShowNewSessionModal(false)}>×</button>
            </div>
            <form
              className="modal-form"
              onSubmit={(e) => {
                e.preventDefault();
                void handleSaveSession();
              }}
            >
            <div className="modal-body">
              <div className="form-group">
                <label>{t.sessionName}</label>
                <input
                  type="text"
                  name="name"
                  value={newSession.name}
                  onChange={handleSessionChange}
                  placeholder={t.inputSessionName}
                />
              </div>
              <div className="form-group">
                <label>{t.hostAddress}</label>
                <input
                  type="text"
                  name="host"
                  value={newSession.host}
                  onChange={handleSessionChange}
                  placeholder={t.inputHost}
                />
              </div>
              <div className="form-group">
                <label>{t.port}</label>
                <input
                  type="number"
                  name="port"
                  value={newSession.port}
                  onChange={handleSessionChange}
                  min="1"
                  max="65535"
                />
              </div>
              <div className="form-group">
                <label>{t.username}</label>
                <input
                  type="text"
                  name="user"
                  value={newSession.user}
                  onChange={handleSessionChange}
                  placeholder={t.inputUsername}
                />
              </div>
              <div className="form-group">
                <label>{t.password}</label>
                <input
                  type="password"
                  name="password"
                  value={newSession.password}
                  onChange={handleSessionChange}
                  placeholder={t.inputPassword}
                />
              </div>
              <div className="form-group">
                <label>{t.group}</label>
                <input
                  type="text"
                  name="group"
                  value={newSession.group}
                  onChange={handleSessionChange}
                  placeholder={t.inputGroup}
                />
              </div>
            </div>
            <div className="modal-footer">
              <button type="button" className="modal-button cancel" onClick={() => setShowNewSessionModal(false)}>{t.cancel}</button>
              <button type="submit" className="modal-button save">{t.save}</button>
            </div>
            </form>
          </div>
        </div>
      )}
    </>
  );
}

export default App;
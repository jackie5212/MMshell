# MMshell（Tauri 2 + Rust + portable-pty + xterm.js）

这是一个单窗口 SSH 终端工具（项目名：MMshell），技术栈如下：

- Tauri 2.0
- Rust 1.77+
- portable-pty 0.8
- xterm.js 5.3

## 功能说明

- 顶部地址栏 + 连接按钮
- 地址格式：`user@host:port`（不写端口时默认 22）
- 后端使用系统 `ssh.exe`，通过 `portable-pty` 启动
- SSH 输出通过 Tauri 事件实时推送到 xterm.js
- 在 xterm.js 里的输入会实时写入 SSH 进程
- 窗口大小变化会自动触发终端大小同步（`resize_ssh`）
- 支持基础快捷键：
  - `Ctrl+C`：有选中内容时复制；无选中内容时发送中断字符（SIGINT）
  - `Ctrl+V`：粘贴到远端终端

## 目录结构

- `src/App.tsx`：前端主界面（地址栏、按钮、xterm、事件绑定）
- `src/App.css`：界面样式
- `src-tauri/src/lib.rs`：Rust 后端（PTY 生命周期与 SSH 命令）
- `src-tauri/Cargo.toml`：Rust 依赖与版本
- `src-tauri/tauri.conf.json`：Tauri 窗口与打包配置

## 运行环境

1. Windows 10/11
2. 安装 [Rust](https://www.rust-lang.org/tools/install)（1.77+）
3. 安装 Node.js 18+
4. 系统里可用 `ssh.exe`（Windows OpenSSH Client）

可用以下命令确认：

```powershell
ssh -V
rustc -V
node -v
```

## 开发运行

在项目根目录 `mmshell-app` 执行：

```powershell
npm install
npm run tauri dev
```

启动后：

1. 在地址栏输入 `user@host:port`
2. 点击“连接”
3. 在终端中按 SSH 标准流程输入密码（密码不会回显）

## 生产打包

```powershell
npm run tauri build
```

产物在：

- `src-tauri/target/release/bundle/`

通常包含：

- `.msi` 安装包
- `.exe` 可执行文件（具体取决于本机打包目标）

## 说明

- 本项目不使用任何第三方 SSH 协议库，仅调用系统 `ssh.exe`。
- 如果连接失败，优先确认：
  - 地址格式是否正确
  - 目标主机端口可达
  - 本机 `ssh.exe` 可直接在 PowerShell 里连接

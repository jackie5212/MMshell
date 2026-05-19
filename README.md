# MMShell

MMShell 是一款基于 **Tauri + Rust** 构建的现代化 SSH/SFTP 客户端，致力于提供比肩 MobaXterm 的强大功能与流畅体验。

## ✨ 核心特性

### 🔐 强大的 SSH 连接
- **多种认证方式**：支持密码认证和私钥认证
- **智能私钥管理**：自动修复私钥文件权限，无需手动配置
- **全格式支持**：支持所有格式的私钥文件（.pem, .key, .ppk 等）
- **兼容性优化**：支持 legacy SSH 服务器和嵌入式设备

### 📁 SFTP 文件传输
- 原生 SFTP 支持
- 与 SSH 会话无缝集成
- 支持私钥认证的 SFTP 连接

### 🎯 会话管理
- **多会话支持**：同时管理多个 SSH/SFTP 会话
- **会话分组**：按组组织和管理会话
- **快速连接**：保存常用会话配置，一键连接
- **会话编辑**：随时修改会话配置

### 🖥️ 终端体验
- 基于 xterm.js 的高性能终端仿真
- 支持 ANSI 颜色和转义序列
- 实时调整终端大小
- 流畅的交互式 Shell 体验

### 🔧 跨平台支持
- Windows（主要支持）
- Linux（计划中）
- macOS（计划中）

## 🚀 快速开始

### 系统要求

- **操作系统**：Windows 10/11
- **SSH 客户端**：需要安装 OpenSSH for Windows（Windows 10+ 已内置）

### 安装

1. 克隆仓库：
```bash
git clone https://github.com/yourusername/MMshell.git
cd MMshell
```

2. 安装依赖：
```bash
cd mmshell-app
npm install
```

3. 启动开发模式：
```bash
npm run tauri dev
```

4. 构建发布版本：
```bash
npm run tauri build
```

## 📖 使用指南

### 创建新会话

1. 点击工具栏的 **"新建会话"** 按钮
2. 填写连接信息：
   - **名称**：会话显示名称
   - **地址**：格式为 `user@host:port`（例如：`root@192.168.0.103:22`）
   - **密码**：登录密码（可选）
   - **私钥文件**：选择私钥文件路径（可选）
   - **分组**：会话所属分组
3. 点击 **"保存并连接"**

### 使用私钥登录

MMShell 会自动处理私钥文件的权限问题：

1. 在新建会话时，点击 **"浏览..."** 选择私钥文件
2. 支持任意路径和格式的私钥文件
3. 系统会在连接前自动修复私钥权限
4. 无需手动执行 chmod 或 icacls 命令

**支持的私钥格式：**
- OpenSSH 格式（id_rsa, id_ed25519 等）
- PEM 格式
- PPK 格式（PuTTY）
- 其他常见格式

### 连接管理

- **断开连接**：点击工具栏的断开按钮
- **切换会话**：从左侧会话列表选择
- **编辑会话**：右键点击会话 → 编辑
- **删除会话**：右键点击会话 → 删除

## 🛠️ 技术栈

### 前端
- **React 18** - UI 框架
- **TypeScript** - 类型安全
- **xterm.js** - 终端仿真
- **Vite** - 构建工具

### 后端
- **Tauri 2.x** - 桌面应用框架
- **Rust** - 系统级编程
- **portable-pty** - 跨平台 PTY 支持
- **Windows API** - 系统集成

### 关键依赖
```toml
tauri = "2"
portable-pty = "0.8"
serde = "1"
windows = "0.61"
```

## 🔒 安全性

### 私钥权限自动修复

Windows OpenSSH 对私钥文件权限要求极其严格。MMShell 实现了智能的权限修复机制：

1. **获取文件所有权** - 确保当前用户拥有文件
2. **重置所有权限** - 清除所有 ACL 条目
3. **移除继承权限** - 禁用权限继承
4. **设置独占权限** - 仅授予当前用户完全控制

这个过程在每次连接前自动执行，对用户完全透明。

### 最佳实践

- 私钥文件建议存放在 `C:\Users\YourName\.ssh\` 目录
- 不要将私钥文件共享给其他用户
- 定期备份重要的私钥文件

## 📋 开发计划

### 已完成 ✅
- [x] 基础 SSH 连接功能
- [x] SFTP 文件传输
- [x] 多会话管理
- [x] 会话分组
- [x] 私钥认证支持
- [x] 自动私钥权限修复
- [x] 支持所有格式的私钥文件

### 进行中 🚧
- [ ] SFTP 图形化文件管理器
- [ ] 端口转发功能
- [ ] SSH 隧道支持

### 计划中 📅
- [ ] 多标签页支持
- [ ] 会话同步（云存储）
- [ ] 主题定制
- [ ] 插件系统
- [ ] Linux/macOS 支持
- [ ] 原生 SSH 实现（不依赖系统 ssh.exe）

## 🤝 贡献指南

欢迎提交 Issue 和 Pull Request！

1. Fork 本仓库
2. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 开启 Pull Request

## 📄 许可证

本项目采用 **Creative Commons Attribution-NonCommercial 4.0 International License** (CC BY-NC 4.0)

您可以：
- ✅ **共享** — 在任何媒介以任何形式复制、发行本作品
- ✅ **演绎** — 修改、转换或以本作品为基础进行创作

但必须遵守：
- 📝 **署名** — 您必须给出适当的信用，提供到本许可证的链接，并标明是否做了修改
- 🚫 **非商业性使用** — 您不得将本作品用于商业目的

详见 [LICENSE](LICENSE) 文件或访问 [CC BY-NC 4.0](https://creativecommons.org/licenses/by-nc/4.0/)

## 🙏 致谢

- [Tauri](https://tauri.app/) - 优秀的桌面应用框架
- [xterm.js](https://xtermjs.org/) - 强大的终端仿真库
- [portable-pty](https://github.com/wez/wezterm/tree/main/pty) - 跨平台 PTY 实现
- [MobaXterm](https://mobaxterm.mobatek.net/) - 灵感来源

## 📞 联系方式

- 项目主页：[GitHub Repository](https://github.com/yourusername/MMshell)
- 问题反馈：[Issues](https://github.com/yourusername/MMshell/issues)

---

**MMShell** - 让远程运维更简单、更高效 🚀

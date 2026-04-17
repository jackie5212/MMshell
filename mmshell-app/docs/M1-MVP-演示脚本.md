# M1 MVP 演示脚本（Day023）

## 演示目标

- 证明 Tauri + Rust + React 主链路已打通
- 演示会话保存、会话列表、连接测试、标签页状态保持
- 演示基础日志可追踪

## 演示步骤

1. 启动应用：`npm.cmd run tauri dev`
2. 观察顶部健康状态 `Rust backend health: ok`
3. 新增会话：输入 `name/host/port/username` 后点击 `Save Session`
4. 验证列表：新会话出现在 `Saved Sessions`
5. 点击 `Reachability Test`，观察连接状态提示
6. 点击 `Open Tab` 打开会话标签页
7. 在终端占位框输入文本，切换标签页后返回，确认内容未丢失
8. 点击 `Copy` 与 `Paste`，验证基础复制粘贴
9. 查看 `Recent Activity Logs`，确认日志新增

## 演示成功标准

- 主流程无崩溃
- 会话可保存与读取
- 标签页切换不丢状态
- 日志中能看到核心操作记录


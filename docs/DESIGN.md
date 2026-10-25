# GitHub Actions 打包阻塞修复

日期：2026-10-25。基线提交：d290fd23b26424292457d39f761882b153d304f4。

## 目标与边界

恢复 Windows/macOS 的 smoke、完整回归、PyInstaller 构建及 artifact 上传。沿用现有 UI、依赖与打包 spec；本次不更改业务规则、不发布 Release。仓库没有 AGENTS.md、AI_CONTEXT.md 或 REQUIREMENT.md，本设计以用户确认的任务与当前配置为依据。

## 已核实的问题

- Actions #108（40182863307）的 Windows job 121137464451 在 smoke 临时目录清理时触发 WinError 32，output.log 的 RotatingFileHandler 仍持有打开的文件。
- Actions #107（40174218570）的 Windows job 121111339449 在并发保存测试中触发 WinError 5：测试在 ZIP 复制过程中替换打开的源文件，不能在 Windows 上按预期到达并发修改检查。源保存实现本身是在关闭 ZIP 句柄后替换。
- 新增同名跨表 Qt 测试调用最后一项的 finish_current，会打开模态完成提示框，没有模拟该调用，offscreen 测试仍会进入嵌套事件循环。#108 macOS 在查询时停留回归步骤；该原因由代码定位，尚未获取已完成的 macOS job 日志证明。

## 修改入口与职责

1. AppLogger.close 明确终止实例，并保持重复调用安全。关闭后的 log 不重开文件。main 在事件循环退出或启动失败收尾时恢复异常钩子、移除 UI sink、关闭日志。
2. smoke 使用 ExitStack 注册服务与日志释放，退出顺序先释放服务、再日志、最后删除临时目录。Qt 测试也明确关闭 logger 并在清理窗口后释放 UI sink。
3. 并发保存测试把外部替换注入到 ZIP 句柄关闭后的 stat 检查前。仍断言外部文件保留、内存未误标、临时文件被清理且重试成功；不跳过 Windows 测试、不放宽异常断言。
4. Qt 自动化测试模拟信息提示框，意外错误框直接使测试失败。业务提示框及 UI 保持原样。
5. workflow 为回归步骤设置 5 分钟、构建 job 设置 30 分钟上限，失败仍阻止打包与发布，避免无人确认的对话框长期占用 runner。

## 验收与剩余风险

- 本地运行 logger 生命周期及轮转回归、Excel 并发故障注入、Qt 跨表与失败重试流程、全部 unittest、smoke、编译和 diff 检查。
- 原生 Windows 文件锁最终需由修复提交上的 Actions 证明；macOS 本地通过不能替代 Windows runner 结果。
- 两个平台均须通过同一修复提交的测试、成品检查、压缩、上传，才能声称远程打包恢复。
- 不把普通分支构建的 release skipped 当作故障；发布仍只由版本标签或显式手动 release_tag 触发。
- 截图权限、F8、原生标注等 UI 验收不在本次自动化修复范围；本次无视觉布局变更。

长期上下文建议：未来建立 AI_CONTEXT.md 时记录 logger 生命周期、自动化测试不能等待模态对话框、Windows 并发文件操作边界与现有打包入口，勿将本次未执行的远程验收写成已通过事实。

## 本轮本地验证

隔离环境 `/private/tmp/s2w-packaging-20261025` 按 requirements 安装独立依赖后，完整 50 项 unittest 通过（含 4 项 Qt offscreen 流程）；服务 smoke、compileall、git diff --check 通过。PyInstaller 6.22.3 / PySide6 6.12.0 的本地 macOS arm64 构建通过，check_bundle 检查通过（284 个文件入口，108.31 MiB 唯一文件字节）。产物位于 `/private/tmp/s2w-packaging-20261025-artifacts`，未创建项目 dist。原生 Windows Actions、远程上传与真实截图验收仍需单独核实，本地构建通过不代表远程已恢复。

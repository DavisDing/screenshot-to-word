# Screenshot To Word

一个面向测试执行场景的跨平台桌面工具：读取 Excel 用例，执行截图与标注，并自动把截图写入 Word 报告。

> 当前主线已重构为 **Python + PySide6 / Qt**，目标是更专业的 UI、更稳定的跨平台能力，以及更强的截图标注体验。旧版 Tkinter 入口已保留在 `/Volumes/DingHao_UP/screenshot-to-word/legacy_tk_main.py`，便于回退对比。

## 核心能力

- **现代化 UI**：Qt 卡片式布局、侧边栏、统一视觉风格，适配 Windows / macOS。
- **Excel 智能列识别**：支持中文、英文、别名和模糊列名匹配；识别不到时会弹出“确认 Excel 列含义”窗口，通过样例值手动说明哪一列是什么。
- **截图标注增强**：支持截图后裁剪、方形标记、圆形标记、荧光笔、画笔、文字、撤销、快捷保存。
- **Word 报告生成**：按测试名称生成或追加 `.docx`，插入验证点、步骤信息和截图。
- **执行回写**：完成当前截图后把 Excel 对应行标记为“已执行”。
- **打包可诊断**：启动异常会写入 `logs/crash.log`，避免打包后“闪退但不知道原因”。

## 项目结构

```text
/Volumes/DingHao_UP/screenshot-to-word
├── main.py                         # PySide6/Qt 新入口
├── legacy_tk_main.py               # 旧版 Tkinter 入口，仅用于回退
├── qt_app/
│   ├── models/test_case.py
│   ├── services/
│   │   ├── excel_service.py        # Excel 预览、智能映射、读取、回写
│   │   ├── word_service.py         # Word 生成/追加
│   │   ├── screenshot_service.py   # Qt 屏幕截图
│   │   ├── logger.py
│   │   └── paths.py
│   └── ui/
│       ├── main_window.py          # 主界面
│       ├── control_window.py       # 执行控制面板
│       ├── annotator_window.py     # 截图标注器
│       ├── excel_mapping_dialog.py # Excel 列映射确认
│       └── style.py                # 全局 QSS 皮肤
├── tests/smoke_test_services.py    # 服务层冒烟测试
├── excel_input/                    # Excel 输入目录
├── word_output/                    # Word 输出目录
└── logs/                           # 日志目录
```

## 安装与运行

建议使用 Python 3.12。

```bash
cd /Volumes/DingHao_UP/screenshot-to-word
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python main.py
```

如果在 macOS 首次截图失败，请到：

```text
系统设置 -> 隐私与安全性 -> 屏幕录制
```

为终端、打包后的 App 或运行环境授予屏幕录制权限。

## Excel 格式

工具不再要求列名完全固定，会自动识别下列语义字段：

| 语义字段 | 是否必要 | 常见列名示例 |
|---|---:|---|
| 测试名称 | 是 | 用例名、用例名称、Case Name、Test Name |
| 验证点 | 是 | 检查点、校验点、Checkpoint |
| 测试结果 | 是 | 执行结果、结果、Status、Result |
| 步骤名称 | 否 | 步骤、操作步骤、Step |
| 步骤描述 | 否 | 步骤说明、操作说明、Desc |
| 预期结果 | 否 | 期望结果、Expected |

当列名不标准时，程序会显示每一列的样例值，让用户手动选择“这一列是什么”。

## 打包

GitHub Actions 已更新为自动构建：

- Windows：`DesktopTestTool.exe`，使用 `--windowed`，不再弹出 CMD 窗口。
- macOS：`DesktopTestTool.app`。
- Release 中会同时上传 Windows 和 macOS zip 包。

本地 Windows 打包示例：

```powershell
pyinstaller main.py --name DesktopTestTool --noconfirm --clean --onefile --windowed --icon=icon.ico --collect-all PySide6 --collect-all shiboken6
```

本地 macOS 打包示例：

```bash
pyinstaller main.py --name DesktopTestTool --noconfirm --clean --windowed --collect-all PySide6 --collect-all shiboken6
```

## 自测

服务层冒烟测试覆盖：Excel 列识别、Excel 回写、Word 插图生成。

```bash
python -m compileall -q main.py qt_app tests
python tests/smoke_test_services.py
```

> 当前环境如果没有安装 PySide6，GUI 不能直接启动；但代码会把缺少依赖等启动问题写入 `logs/crash.log`。安装 `/Volumes/DingHao_UP/screenshot-to-word/requirements.txt` 后即可运行 Qt 界面。

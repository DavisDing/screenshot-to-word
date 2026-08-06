# Screenshot To Word

一款面向测试执行场景的跨平台桌面工具：读取 Excel 用例，按步骤执行截图与标注，并自动生成 Word 报告，同时回写 Excel 执行状态。

当前主线已经重构为 **Python + PySide6 / Qt**，目标是：

- 更现代、更专业的 UI
- 更稳定的跨平台体验（Windows / macOS）
- 更顺畅的步骤式测试执行流程
- 更强的截图标注能力
- 更可靠的打包发布与版本管理

## 当前最新需求 / 目标

### 1. UI 方向
- 采用更现代的界面皮肤
- 主界面布局更清爽、更专业
- 执行控制面板更轻量，尽量不遮挡测试内容
- Excel 列映射对话框更简洁，减少视觉占用
- 软件内展示版本号，打包版显示 Release 版本

### 2. 截图标注方向
- 支持截图后标注
- 逐步增强标注能力，例如：
  - 裁剪
  - 正方形 / 矩形标记
  - 荧光笔
  - 文字标注
  - 撤销 / 重做
- 截图时主窗口与执行面板需要正确隐藏，避免进入截图画面

### 3. Excel 执行方向
- 自动识别 Excel 列含义
- 识别不到时提供人工确认，尤其要能看样例值判断“这一列是什么”
- 步骤版用例必须严格按步骤逐条执行
- 合并单元格场景要能正确读取测试名称、验证点
- 执行完成后回写 Excel 状态
- 支持“下一条 / 跳过当前案例”等执行动作

### 4. Word 输出方向
- 自动按用例生成或追加 Word
- 验证点不重复输出
- 跳过案例也要能同步记录到 Word
- 输出内容尽量结构清晰、适合测试报告场景

### 5. 打包与版本方向
- GitHub Actions 自动打包 Windows / macOS
- 打包后软件内要显示 Release 版本号
- 本地开发显示“开发版”
- 尽量避免 Windows 打包后弹出 cmd 窗口闪退问题

## 技术栈

- 语言：Python
- GUI：PySide6 / Qt
- Excel：openpyxl
- Word：python-docx
- 图像：Pillow
- 打包：PyInstaller
- CI/CD：GitHub Actions

要求的 Python 版本以当前依赖和 CI 为准，当前仓库使用 Python 3.12 进行构建验证。

## 现有主线结构

```text
/Volumes/DingHao_UP/screenshot-to-word
├── main.py                         # 程序入口
├── qt_app/
│   ├── version.py                  # 版本读取
│   ├── models/test_case.py         # 测试项数据模型
│   ├── services/
│   │   ├── excel_service.py        # Excel 读取、映射、回写
│   │   ├── word_service.py         # Word 生成/追加
│   │   ├── screenshot_service.py   # 截图服务
│   │   ├── logger.py               # 日志
│   │   └── paths.py                # 路径与目录管理
│   └── ui/
│       ├── main_window.py          # 主窗口
│       ├── control_window.py       # 执行控制面板
│       ├── annotator_window.py     # 截图标注窗口
│       ├── excel_mapping_dialog.py # Excel 列确认对话框
│       ├── icons.py                # 图标加载
│       └── style.py                # Qt 样式
├── scripts/
│   ├── make_icns.py                # macOS 图标生成
│   └── write_build_version.py      # 打包版本写入
├── tests/smoke_test_services.py    # 冒烟测试
├── requirements.txt                # 依赖
├── icon.ico                        # Windows 图标
├── assets/                         # Qt 图标资源
├── excel_input/                    # Excel 输入目录
├── word_output/                    # Word 输出目录
└── logs/                           # 日志目录
```

## 运行方式

```bash
cd /Volumes/DingHao_UP/screenshot-to-word
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python main.py
```

## 自测命令

```bash
python -m compileall -q main.py qt_app tests scripts
python tests/smoke_test_services.py
```

## 打包说明

- Windows：PyInstaller 打包为 `.exe`，使用 `--windowed`
- macOS：PyInstaller 打包为 `.app`
- GitHub Actions 会在打包前写入版本号
- 打包版软件内部会显示 GitHub Release 的 tag

## 已完成里程碑

- 已完成 Qt 主线重构
- 已完成主界面现代化
- 已完成执行控制面板
- 已完成 Excel 列映射确认流程
- 已完成步骤版合并单元格支持
- 已完成 Word 验证点去重
- 已完成打包版本号注入
- 已完成基础 smoke test

## 后续优先事项

1. 继续完善“跳过当前案例”逻辑，并同步写入 Word / Excel
2. 继续优化截图标注工具能力
3. 继续收敛执行控制面板与 Excel 映射对话框的 UI
4. 增加更完整的 GUI 回归测试
5. 验证 GitHub Actions 产物在 Windows / macOS 的实际可用性

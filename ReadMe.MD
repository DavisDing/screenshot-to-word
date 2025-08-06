以下是为该桌面自动化测试工具编写的完整 README.md 文件，适合用于开源发布、协作开发或内部交付：

⸻


# 桌面自动化测试工具（DesktopTestTool）

## 🧭 工具简介

本工具旨在提升手工测试效率，适用于在执行测试用例时快速截图、标注、记录验证点，并生成带图文说明的 Word 报告。适配 Windows 操作系统，可打包为 `.exe` 独立运行。

---

## ✨ 功能特点

- ✅ 读取指定目录的测试用例 Excel 文件（测试名称、验证点、执行状态）
- ✏️ 截图并支持画圈、加文字标注
- 📄 自动生成图文并茂的 Word 报告
- ✅ 每执行一条用例即自动更新 Excel 状态列
- ⚡ 快捷键 F8 或按钮操作
- 💡 支持跳过、退出操作
- 📋 自动日志记录，保留执行历史

---

## 📁 项目结构

DesktopTestTool/
├── main.py                 # 启动入口，主窗口逻辑
├── core/
│   ├── test_runner.py      # 执行控制流程
│   └── screenshot.py       # 截图模块（pyautogui）
├── ui/
│   ├── control_panel.py    # 控制面板 UI（置顶 + 快捷操作）
│   └── annotator.py        # 图像标注界面
├── utils/
│   ├── excel_handler.py    # Excel 文件读写处理
│   └── logger.py           # 日志记录器
├── excel_input/            # 运行时读取的用例 Excel 文件
├── word_output/            # 输出 Word 文档目录
├── logs/                   # 日志文件目录
├── requirements.txt        # 第三方依赖列表
└── README.md               # 工具说明文档

---

## ✅ Excel 文件格式说明

| 文件名 | 验证点 | 执行结果 |
|--------|--------|----------|
| Login  | 登录成功后跳转首页 | (留空或非“已执行/PASS”) |
| ...    | ...    | ...      |

- 工具会自动跳过 **已执行、PASS、Passed** 的行
- 每次截图后自动更新状态为“已执行”

---

## 🚀 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt

2. 准备目录结构

确保 main.py 所在目录下存在以下目录：

excel_input/     # 放入 Excel 文件
word_output/     # Word 报告输出目录
logs/            # 自动生成执行日志

3. 运行工具

python main.py


⸻

💻 打包为 .exe（可选）

需安装 pyinstaller：

pip install pyinstaller
pyinstaller main.py --onefile

生成后的 dist/DesktopTestTool.exe 可独立运行，但需确保与以下文件夹同级：

DesktopTestTool.exe
├── excel_input/
├── word_output/
├── logs/


⸻

🧪 功能测试建议清单

测试点	说明
Excel 文件缺失	应弹窗提示
Excel 列缺失	明确提示缺“文件名/验证点/执行结果”
Excel 被打开中	写入失败时应给出提示
Word 被打开中	插图失败不应中断程序
ESC 退出标注	应保存标注或跳过
多张截图	应正确编号保存，不覆盖
F8 快捷键	与按钮响应一致


⸻

📌 依赖说明（requirements.txt）

pandas
openpyxl
python-docx
pyautogui
keyboard
Pillow


⸻

📞 联系与反馈

如有功能建议或问题反馈，欢迎提交 Issue 或联系开发者。

---

是否还需要 `.spec` 打包配置，或生成 GitHub Actions 自动构建？我可以继续补充。
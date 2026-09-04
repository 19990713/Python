# CLAUDE.md

本文件为 Claude Code（claude.ai/code）在本仓库中工作时提供指导。

## 项目概述

印特 (YinTe) 自动化 —— 针对 Windows 桌面打印店管理应用的 Python GUI 自动化。读取 Excel 或 PDF 输入，并自动填写「新建工单」DataGridView 表单。

这里有两个独立项目：
- **`yinte-automation/`** —— 主自动化套件（入口、核心引擎、PDF 分析器、Excel 读取器）
- **`图纸统计.py`** —— 独立的 PDF 页面尺寸统计工具（PDF 分析器模块的前身）

## 命令

```bash
# 本地运行（把文件拖到批处理脚本上，或通过命令行）：
py -3.12 "d:\CODE\yinte-automation\entry.py" "文件.xlsx"      # Excel 模式
py -3.12 "d:\CODE\yinte-automation\entry.py" "图纸.pdf"       # PDF 模式
py -3.12 "d:\CODE\yinte-automation\entry.py" "文件夹/"         # PDF 文件夹

# 仅 Excel 的入口（需要先手动打开「新建工单」窗口）：
py -3.12 "d:\CODE\yinte-automation\main.py" --excel "数据.xlsx"
py -3.12 "d:\CODE\yinte-automation\main.py" --excel "数据.xlsx" --dry-run

# 打包单文件 EXE：
cd d:\CODE\yinte-automation
py -3.12 -m PyInstaller 印特自动录入.spec

# 检查运行中窗口的 GUI 控件：
py -3.12 "d:\CODE\yinte-automation\inspector.py" --title "新建工单"
```

桌面批处理文件位于 `D:\Users\桌面\印特自动录入.bat`。打包后的 EXE 同时位于 `d:\CODE\yinte-automation\dist\印特自动录入.exe` 和 `D:\Users\桌面\印特自动录入.exe`。`config.yaml` 作为兜底默认值被嵌入 EXE —— 只有在需要覆盖某些值时，才在 EXE 旁边放置外部 `config.yaml`。

## 依赖

运行时依赖（声明于 `yinte-automation/requirements.txt`）：

- `pywinauto` —— GUI 自动化（查找窗口、点击、键盘）
- `pyperclip` —— 用于粘贴产品代码的剪贴板
- `pypdf` —— PDF 页面尺寸提取（`src/pdf_analyzer.py`）
- `pandas` —— Excel 批量读取（`src/excel_reader.py`）
- `openpyxl` —— 轻量单单元格读取（B4）+ 生成统计 Excel
- `pyyaml` —— 加载 `config.yaml`
- `uiautomation` —— 补录模式下 NewItem 行定位（`click_new_row`，惰性导入）

仅用于构建（运行时不会导入）：`PyInstaller` —— 打包单文件 EXE。

注意：`requirements.txt` 曾一度遗漏 `uiautomation`；全新安装会在补录模式下崩溃。请保持它被列出。

## 架构

### 入口

- **`entry.py`** —— 统一的拖放入口。自动识别输入类型（Excel 还是 PDF/文件夹），并通过双击印特主窗口相对位置 (128, 79) 处的按钮**自动打开「新建工单」窗口**。这是面向用户的主要入口。
- **`main.py`** —— 仅 Excel 的 CLI 入口。要求「新建工单」窗口已经打开。使用 argparse，带 `--excel`、`--dry-run`、`--start-row` 参数。

### 核心引擎（`src/automation.py`）

印特应用使用 Windows Forms DataGridView。塑造设计的关键约束：

1. **动态控件 ID** —— 每次会话都会生成新的控件 ID。必须使用标签文本 + 相对位置来定位，绝不能使用 `control_id`。
2. **DataGridView 单元格没有常驻的 Edit 控件** —— Edit 控件只在单元格正在被编辑时才出现。必须使用鼠标点击 + 键盘导航，而不是 `set_text()` 或控件操作。
3. **使用剪贴板粘贴以提速** —— 使用 `pyperclip.copy()` + `keyboard.send_keys("^v")`，而不是逐字符键入。这是最快的方式。
4. **每一行的网格导航流程**：
   - 第 1 行：点击坐标 (window.left + 111, window.top + 375) 处的单元格 → 粘贴代码 → Enter → Enter（跳过实价/价格）→ 粘贴份数 → Enter → 粘贴页数 → Enter
   - 第 N 行：Enter（跳到下一行的编号列）→ 相同的填写序列
   - 最后一行之后：额外一次 Enter 来完成
5. **win32 后端** —— 64 位 Python 控制 32 位印特。`win32` 后端可用；`uia` 对本应用不可靠。

### PDF 分析器（`src/pdf_analyzer.py`）

通过 `pypdf` 提取 PDF 页面尺寸，匹配标准尺寸（A0-A5、B0-B5），并生成尺寸统计。`generate_excel()` 函数存在，但在直接模式下**并未使用** —— 拖入 PDF 时，记录直接传给引擎，不经过中间 Excel。

### Excel 读取器（`src/excel_reader.py`）

读取 `.xlsx/.xls`，表头行可配置。关键细节：`read_cell()` 把诸如 "B4" 的单元格地址解析为行/列索引（1 起始 → 0 起始转换）。子项类型默认从 B4 单元格读取。

### 配置文件（`config.yaml`）

外部 YAML，必须与 EXE/批处理文件放在一起。包含：
- 窗口定位（`rel_x: 111`、`rel_y_base: 375`、`row_height: 24`）
- 30 个产品代码，键为 `子项_尺寸分类` → 数字代码（例如 `"蓝图_A1+": "200043"`）
- Excel 设置（`header_row: 6`、`subtype_cell: "B4"`）
- 列名映射

### EXE 打包

通过 `印特自动录入.spec` 使用 PyInstaller `--onefile`。EXE 是控制台应用（非 windowed），因此 print/input 提示可见。运行时，`entry.py` 通过 `sys.frozen` 检测冻结模式，并把 `APP_DIR` 解析为 `sys.executable.parent`，使 config/logs 相对于 EXE 而非临时目录。

## 关键陷阱

- **绝不要触发表单的默认「确定」按钮** —— 在表单本身上（而不是网格单元格内）按 Enter 会关闭窗口。`_fill_row()` 以页数 + Enter 结束；下一行以 `_fill_row()` 开始，其开头是 `^v`（粘贴），而不是 Enter。
- **最后一行会额外按一次 Enter** 来完成，但仅当至少有一行成功时。这是安全的，因为在 `_fill_row` 之后焦点仍在 DataGridView 中。
- **子项类型在 B4 单元格，而不是 B6** —— 早期曾混淆；B4 才是正确的单元格。
- **产品代码键格式** —— `{子项}_{尺寸分类}`，例如 `"蓝图_A1+"`。代码查找回退到裸 `{尺寸分类}`，再回退到裸 `{子项}`（用于不随尺寸变化的代码，如硫酸纸）。
- **打开「新建工单」使用类名回退** —— 当标题匹配失败时（win32 后端的中文编码问题），主印特窗口通过 `class_name_re="WindowsForms10.*Window\.8\.app.*"` 查找。
- **记录在录入前按 A4→A3→A2→A2+→A1→A1+→A0→A0+ 排序**。
- **录入成功时自动关闭** —— 仅在出错或部分失败时保持控制台窗口打开。
- **需要 Python 3.12** —— 更新版本（3.14）与 pandas 存在包兼容性问题。
- **中文编码** —— 批处理文件必须是纯 ASCII，以避免在 cmd.exe 中出现乱码。
- **绝不要在控制台输出中使用 emoji** —— Windows GBK 控制台无法编码 emoji（`📄` `⚠️` `❌` `📋` `📌` `⏳`），因此任何包含它们的 `print()` 都会抛出 `UnicodeEncodeError` 并立即崩溃应用（闪退）。它们已从 `entry.py`/`pdf_analyzer.py` 中移除；不要重新引入。纯中文文本和 `→`/`—` 是 GBK 安全的，没有问题。
- **尺寸加长后缀始终用 `"+"`，不用 `"加长"`** —— 尺寸名称如 `"A1+"`，产品代码键名如 `"蓝图_A1+"`。PDF 分析器、Excel 导入、排序、配置等所有地方统一用 `"+"`。旧 Excel 中 `"A1加长"` 在导入时自动转为 `"A1+"`。**绝对不要改回中文。**

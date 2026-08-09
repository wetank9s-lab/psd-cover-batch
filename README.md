# PSD 批量封面生成器（Excel 数据驱动）

用一张 PSD 模板 + 一份 Excel 名单，自动把**姓名 / 电话 / 销售顾问**填进对应文字图层，
并按**门店**自动切换对应的门店 Logo，批量导出成一张张 PNG 封面图。

适合抖音 / 视频号等渠道的「活动海报 / 视频封面」批量制作：
一张模板，N 个门店销售顾问，每人一张带自己姓名、电话、门店 Logo 的封面，全程无需手动打开 Photoshop 改字。

> **统一渲染内核（Stage 5）**：Preview（单张预览）、Batch（批量生成）、CLI（命令行）
> 三条入口共用 `core/renderer.py` 的**唯一单行渲染入口 `render_one()`**，
> 文字写入后做 read-back 校验、Logo 可见性写入后回读验证、导出后校验文件存在且非空，
> 单行失败不中断批次，全部汇总进 `BatchResult`。核心逻辑可脱离 Photoshop 单测（fake session）。

---

## 效果示例

假设 PSD 模板里有这些图层：

```
电话信息 (组)
  ├─ 姓 名        ← 姓名文字层
  ├─ 电话 拷贝     ← 电话文字层
  └─ 销售顾问 拷贝 ← 职位文字层
易田电器          ← 门店 Logo 图层（每层对应一个门店）
欣盛电器
华林电器
...
```

Excel 每行一人：

| 门店(A) | 姓名(B) | 销售顾问(C) | 电话(D) |
| --- | --- | --- | --- |
| 易田电器 | 刘超 | 销售顾问 | 137****3719 |
| 华林电器 | 王兵 | 销售顾问 | 180****2479 |

程序会为每行生成一张 PNG：填好姓名电话、只显示该门店的 Logo、其余门店 Logo 隐藏。

---

## 项目结构

```
psd-cover-batch/
├── qifang_cover_maker.py      # 图形界面版（GUI，tkinter）
├── psd_cover_batch.py         # 命令行版（CLI）
├── core/                      # 可单测的核心逻辑（不依赖 Photoshop / tkinter）
│   ├── renderer.py            #   Stage 5 统一渲染内核：render_one() / run_batch()
│   ├── photoshop.py           #   Photoshop COM 会话管理（打开/关闭/暂存盘清理）
│   ├── layer_index.py         #   图层索引 + LayerRef 逐层 resolve（禁按名 find）
│   ├── logo_mapping.py        #   门店 Logo / 品牌 Logo 映射（模糊包含匹配）
│   ├── excel_data.py          #   Excel 解析：ExcelRow / ExcelDataset / 数据清洗
│   ├── output_paths.py        #   输出目录：GroupFolderMap / 分组 / 碰撞处理
│   └── util.py                #   文件名清洗等
├── tests/                     # pytest 测试（422 个用例，全部可离线跑）
└── docs/GUI_WORKFLOW.md       # GUI 操作流程基线
```

## 环境依赖

- **Windows** + 已安装 **Adobe Photoshop**（通过 COM 自动化驱动，脚本会调用本机 Photoshop）
- Python 3.8+
- 依赖库：`pywin32`、`openpyxl`

> 说明：脚本依赖 Photoshop 的 COM 接口，因此**必须在装了 Photoshop 的 Windows 上运行**。
> 纯命令行服务器（Linux / macOS）环境无法直接运行。

---

## 安装

```bash
pip install -r requirements.txt
```

---

## 快速开始

### 1. 先看模板图层（确认图层名）

```bash
python psd_cover_batch.py --psd 8081.psd --inspect
```

只查 PSD 图层树时**无需 Excel**。提供 `--xlsx` 时会额外输出「匹配到的门店 Logo 图层」
（评分制 `match_store_logo`，歧义不自动选）与「品牌 Logo 图层」建议：

```bash
python psd_cover_batch.py --psd 8081.psd --xlsx 门店销售0803销售顾问.xlsx --inspect
```

如果名字对不上，按下面「PSD 模板规范」改脚本顶部的映射常量即可。

### 2. 先合成一行做测试

```bash
python psd_cover_batch.py --psd 8081.psd --xlsx 门店销售0803销售顾问.xlsx --row 1 --out ./out
```

只生成 Excel 第 1 行对应的 PNG，用于肉眼检查排版 / 字体 / Logo 是否正确。

### 3. 批量生成

```bash
python psd_cover_batch.py --psd 8081.psd --xlsx 门店销售0803销售顾问.xlsx --out ./out
```

每一行数据导出一张 PNG，命名格式：`001_易田电器_刘超.png`（序号_门店_姓名）。

### 4.（可选）按 Excel 某列分组输出

```bash
python psd_cover_batch.py --psd 8081.psd --xlsx 门店销售0803销售顾问.xlsx \
  --out ./out --group-output-column A
```

按指定列（支持列字母如 `A`/`AA`，或 0-based 数字列号）的值创建子文件夹，例如 `out/易田电器/001_易田电器_刘超.png`：

- 列值为空的行归入 `未分类/` 子文件夹；
- 不同门店名清洗后同名（如 `A/B` 与 `A\B`）自动追加 `_2`、`_3` 等后缀防碰撞；
- 分组目录与输出文件名一样经过路径安全校验（防越界写入）。

---

## 图形界面版（GUI）

除命令行脚本外，仓库还附带带图形界面的「七方视频封面批量制作」程序（`qifang_cover_maker.py`），
功能与命令行版一致（模糊包含匹配门店 Logo + 品牌 Logo 常显 + 姓名不缩放 + 按列分组输出），
但用窗口勾选配置，更适合日常使用：

- 三个标签页：**文件与字段 / 门店 Logo 映射 / 生成**
- 自动识别门店 Logo（模糊包含匹配）与品牌 Logo（名字含 `logo` 的图层每张都显示）
- 支持**按 Excel 任意列分组输出**（下拉选列），生成前做 preflight 校验（列是否有效、路径安全、碰撞统计）
- 预览（单张）与批量生成共用同一渲染内核，预览输出到 `_preview/` 子目录，不污染正式输出
- 配置（门店→Logo 映射、文字图层名等）自动保存到同目录 `qifang_cover_config.json`

### 直接用 exe（推荐）

从 GitHub Releases 下载 `七方视频封面批量制作.exe`，双击即可使用（仍需本机安装 Photoshop）。
> exe 为单文件 windowed 版（无控制台窗口）；因 GUI 无日志界面，若遇问题可改用源码运行看完整报错。

### 从源码运行 / 打包（注意 tkinter）

GUI 依赖 `tkinter`，**打包或运行所用的 Python 必须自带 tkinter**（如系统 Python 3.14）。
托管 Python 3.13 默认不含 tkinter，会报 `ModuleNotFoundError: No module named 'tkinter'`。

```bash
# 用系统 Python 3.14 建 venv（自带 tkinter 8.6）
python3.14 -m venv venv314
venv314\Scripts\python.exe -m pip install --no-cache-dir pywin32 openpyxl pyinstaller

# 直接跑源码
venv314\Scripts\python.exe qifang_cover_maker.py

# 打包成单文件 exe（windowed 无控制台）
venv314\Scripts\python.exe -m PyInstaller --onefile --windowed --name "七方视频封面批量制作" ^
  --hidden-import win32com --hidden-import win32com.client --hidden-import pythoncom ^
  --hidden-import openpyxl qifang_cover_maker.py
```

> 安装依赖时**不要** `pip install --upgrade pip`，否则部分沙箱环境会因删除旧包触发安全拦截导致失败；用 `--no-cache-dir` 即可。
>
> 注意：GUI 与 CLI 都依赖 `core/` 包（可测试纯函数）。**打包/运行时请保持仓库目录结构**
> （`qifang_cover_maker.py` 与 `core/` 同级），PyInstaller 会自动收集 `core`。

---

## 开发 / 测试

纯函数（文件名清洗、Excel 解析、门店/图层名匹配、输出分组、统一渲染内核）位于 `core/`，
**不依赖 Photoshop 与 tkinter**，可直接用 pytest 测试：

```bash
pip install pytest
pytest tests/
```

当前测试：**422 个用例**（`tests/test_cli.py`、`test_excel_dataset.py`、`test_excel_parsing.py`、
`test_filename.py`、`test_group_output.py`、`test_gui_state.py`、`test_gui_styles.py`、
`test_gui_view_model.py`、`test_layer_index.py`、`test_logo_mapping.py`、`test_photoshop_session.py`、
`test_renderer.py`、`test_worker_events.py`），全部可离线运行。
`test_cli.py` 额外做源码级红线扫描（AST 去 docstring 后校验 CLI 不含 Excel 解析、
`Document.Duplicate`、`TextItem.Contents`、`Visible=`、`SaveAs` 等业务实现）。
GUI 操作流程基线见 `docs/GUI_WORKFLOW.md`。

---

## 参数说明

| 参数 | 必填 | 说明 |
| --- | --- | --- |
| `--psd` | 是 | PSD 模板路径 |
| `--xlsx` | 是 | Excel 数据路径（第一行为表头） |
| `--out` | 否 | PNG 输出目录，默认 `./out` |
| `--row` | 否 | 只合成指定行（从 1 开始）；不填则全部 |
| `--group-output-column` | 否 | 按 Excel 该列的值创建输出子文件夹（支持列字母如 `AA` 或 0-based 数字如 `26`）；不指定则不分组 |
| `--brand-logo` | 否 | 品牌 Logo 图层名（每张封面都强制显示）。可重复指定；不指定则自动识别「名字含 logo 且非门店名」的图层 |
| `--inspect` | 否 | 只打印图层树后退出，不做任何导出 |

---

## PSD 模板规范

脚本顶部有可直接修改的映射常量（对应 `psd_cover_batch.py` 开头）：

```python
NAME_LAYER = "姓 名 拷贝"      # 姓名文字图层
PHONE_LAYER = "电话 拷贝"       # 电话文字图层
TITLE_LAYER = "销售顾问 拷贝"   # 职位 / 头衔文字图层
```

- 文字图层按**图层名**精确匹配（注意模板里名字可能带空格，如 `姓 名`）。
- 门店 Logo 图层：程序会自动把「**顶层图层名与 Excel 门店名「模糊包含」匹配**（互含即可，
  不要求完全一致）」的图层当作门店 Logo。例如 Excel 里写 `康乐`，PSD 里图层叫 `康乐电器`，
  也能正确匹配；`九兴` ↔ `九兴电器`、`诚信` ↔ `诚信电器` 同理。
  处理某一行时，**只显示该门店对应的 Logo 图层，隐藏其它门店 Logo**。
- **品牌 Logo（每张都显示）**：名字里包含 `logo`（且不正好等于某个门店名）的图层，
  会被当作「品牌 Logo」，在每一张封面里**强制显示、不会被隐藏**。典型例子：模板里的
  `七方logo`、`七方logo 拷贝` 这类固定品牌角标。若自动识别不准，可用 `--brand-logo 七方logo`
  显式指定（可重复多次）。
- 姓名不做字号缩放：2 字、3 字、4 字姓名统一按模板原始字号写入，保持版式一致。
- **文字写入 / Logo 可见性写入后都会回读验证**，失败即报错并停止该行导出（不会静默导出错误图）。

若你的模板图层命名不同，改上面三个常量、或调整 Excel 列索引即可：

```python
STORE_COL, NAME_COL, TITLE_COL, PHONE_COL = 0, 1, 2, 3  # A,B,C,D 列
```

---

## Excel 表格规范

- 第一个工作表（Sheet1）的第一行为表头，第二行起为数据。
- 列顺序（可在脚本里改索引）：
  - **A 列 = 门店**（用于匹配门店 Logo 图层名）
  - **B 列 = 姓名**
  - **C 列 = 销售顾问 / 职位**
  - **D 列 = 电话**（数字或文本均可，会自动转成字符串）
- 门店名必须与 PSD 里的门店 Logo 图层名**存在包含关系**（互含即可，见上方「PSD 模板规范」）。
- **电话列**支持数字或文本，会自动转成字符串；空姓名 / 空电话的行会被自动跳过并记录原因。

---

## 渲染可靠性保障（Stage 5 统一渲染内核）

`core/renderer.py` 是所有导出路径的**唯一单行渲染入口**，对每一行保证：

1. **Duplicate → 改副本 → 导出 → 关闭**：每一行都复制模板后修改，绝不复用同一文档连续改多行；
   无论文字 / Logo / 导出 / 校验哪一步失败，副本文档都会被 `finally` 关闭（不泄漏 Photoshop 文档）。
2. **文字 read-back 校验**：写入后回读实际图层文字，不一致即报错（`TextVerificationError`），
   不会静默导出错字图。
3. **Logo 可见性回读验证**：按门店切换 Logo 可见性后回读确认，失败即阻止该行导出。
4. **导出后验证**：检查文件存在且大小 > 0（`ExportVerificationError`）；
   多格式导出（如 PSD + 另存 PNG）部分失败时，**保留已成功的文件**，并在错误中列出。
5. **单行失败不中断批次**：失败行记入 `BatchResult`，其余行继续；支持取消（stop 后剩余行标记 CANCELLED）。
6. **COM 跨线程安全**：`SaveOptions` 等 COM 对象**每次导出新建**，绝不模块级缓存
   （COM 对象绑定创建线程的 STA 单元，跨线程复用会报 `CDispatch can not be converted to a COM VARIANT`）。

---

## 常见问题

**Q：运行报错 `module 'pythoncom' has no attribute ...` 或 COM 相关错误？**
A：确认本机已安装 Photoshop，且 `pip install pywin32` 成功。脚本已内置「应用程序正忙」自动重试。

**Q：导出的图里门店 Logo 没切换 / 多显示了别的门店？**
A：用 `--inspect` 确认「匹配到的门店 Logo 图层」列表。门店 Logo 采用「模糊包含」匹配
（如 Excel `康乐` 可匹配 PSD `康乐电器`），一般无需逐字一致；若仍匹配不到，
请修改 PSD 图层名或 Excel 门店名使二者存在包含关系。

**Q：模板里的品牌 Logo（如 `七方logo`）在导出图里不显示？**
A：品牌 Logo 图层必须在模板里处于「显示」状态，且名字含 `logo`（或不等于门店名）。
脚本会强制把它设为可见；若你的图层名不含 `logo`，用 `--brand-logo 七方logo` 显式指定即可。

**Q：用完后 Photoshop 一直占着内存 / 暂存盘（C 盘被吃满）？**
A：旧版会拉起 Photoshop 但从不退出，进程常驻、暂存文件不释放。现已修复：
每次「加载 / 预览 / 生成」结束后会**关闭所有文档并退出由本程序拉起的 Photoshop**，
释放内存与暂存盘；只有**你自己在用** Photoshop 时（程序启动前就已打开），本程序才不打扰、
不会退出它。程序窗口关闭时也会再做一次清理。
注意：PNG 是按当前文档状态导出的，若图层本身被锁定/隐藏且 PS 不允许修改可见性，需先在模板里解锁。

**Q：文字没替换成功？**
A：用 `--inspect` 看实际文字图层名，把 `NAME_LAYER / PHONE_LAYER / TITLE_LAYER` 改成模板里的真实名字。
文字写入后程序会回读验证，若回读与写入不一致会明确报 `TextVerificationError`，请检查模板该图层是否被锁定 / 是否可编辑。

**Q：批量生成时有的行失败，其它行还正常吗？**
A：正常。统一渲染内核单行失败不中断批次，失败的行走入 `BatchResult` 明细（含行号、字段、原因），
其余行继续生成；多格式导出部分失败时，已成功的文件会保留。

**Q：Preview 正常但 Batch 全部报 `Objects of type 'CDispatch' can not be converted to a COM VARIANT`？**
A：这是 COM 对象跨线程复用问题（旧版缺陷）：Preview 在主线程、Batch 在后台线程，复用了绑定主线程的
`SaveOptions` 对象。Stage 5 已修复——每次导出都新建 `SaveOptions`，不再做模块级缓存。请使用新版本（>= v1.1.0）。

**Q：能导出 JPG / 多页 PDF 吗？**
A：当前默认导出 PNG（`PNGSaveOptions`）。如需其它格式，可在 `run()` 里把
`Photoshop.PNGSaveOptions` 换成 `Photoshop.JPEGSaveOptions` 等并调整 `SaveAs` 后缀。

---

## 许可证

MIT —— 可自由用于商业 / 非商业用途。

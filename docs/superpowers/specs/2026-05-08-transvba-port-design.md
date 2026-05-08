# TransVBA → Python 移植设计

**日期**:2026-05-08
**作者**:基于用户与 Claude 的交互式 brainstorm
**目标输出**:Python 桌面工具(GUI),功能 100% 等同于现有 VBA 工程 `D:\Code2Syn\TransVBA\`

---

## 1. 目标与范围

### 1.1 一句话目标

把 `D:\Code2Syn\TransVBA\` 下的 VBA Word 格式刷新插件,**完整、一比一**地移植成 Python 桌面应用,以开源项目 [`cwyalpha/Word-Formatter-Pro`](https://github.com/cwyalpha/Word-Formatter-Pro)(下称 WFP)的代码组织/工程能力为底座。

### 1.2 In Scope

- VBA 工程的**全部**功能特性,一个都不能漏:
  - 5 级标题识别(数字编号 `1`/`1.1`/`1.1.2`/...)与每级独立的字体/字号/对齐/段前段后/行距/加粗
  - 正文格式(字体/字号/行距/对齐/缩进/特殊缩进/段前段后行数)
  - 多级列表自动编号读取与级别识别
  - 目录(TOC)识别与样式(TOC1/TOC2/TOC3 + 手写"Tab + 页码"两种)
  - 表格(题注 + 表体)与图片题注
  - ASCII / 数字字体统一为 Times New Roman
  - 设置持久化("记忆本次设置")
  - 自动识别开关(`AutoDetectNumericTitles` / `AutoDetectIncludeListParagraphs`)
- WFP 的代码组织风格(扁平 `xxx_*.py` 命名)与 python-docx + pywin32 双层 API 策略
- 全新设计的 Tkinter UI(VBA UserForm 的等价替代,但布局现代化)
- TDD 驱动的实现流程

### 1.3 Out of Scope(明确砍掉)

- WFP 的领域逻辑:公文标题检测(`一、`/`（一）`)、附件分页、Markdown 清理、标点规范化、副标题处理、页面布局(页边距/页码/页脚距离)、特定公文字体预设
- CLI 入口、Agent Skill bundle、批处理拖拽、文件队列管理
- VBA Registry 导入兼容层
- 跨平台支持(Linux/macOS):VBA 本身就是 Windows + Office 路径,移植后保持同样要求

---

## 2. 战略决策(已与用户确认)

| 决策点 | 决定 |
|---|---|
| 整合策略 | 只取 WFP 的**架构**,WFP 的领域逻辑全部移除;VBA 的全部功能在 WFP 骨架里用 Python 重写 |
| API 路径 | **python-docx 主导 + pywin32 COM 兜底**(路径 A):大部分操作走 python-docx + 自写 OOXML helpers;唯独多级列表渲染编号读取走 COM(行为与 VBA 100% 一致) |
| 平台 | Windows + 已安装 Word 或 WPS Office |
| UI | 全新设计(左侧分类树 + 右侧详情面板),不 1:1 复刻 UserForm |
| 文件模式 | 单文件模式(每次一个 .docx,等同 VBA `ActiveDocument`),无批处理 |
| 设置持久化 | JSON 单一来源,无 Registry 双写、无 Registry 导入 |
| TDD | 严格执行,Red→Green→Refactor;测试覆盖率门槛见 §9 |

---

## 3. 项目结构

沿用 WFP 的扁平 `xxx_*.py` 风格,因为 VBA `FormatModule.bas` 1500 行单文件移植后过长不利可读,把 core **横向拆**成多个文件,**仍然全部在仓库根目录**(不引入子包)。

```
tvba.py                       ← Tkinter GUI 入口(对应 VBA 的 ShowFormatSettings)
tvba_gui.py                   ← UI 视图(View)
tvba_settings.py              ← FormatSettings 数据类(对应 Typemodule.bas)
tvba_persistence.py           ← JSON 设置读写
tvba_config.py                ← 默认值表(集中存放)
tvba_utils.py                 ← 字号字符串↔点数 / 厘米↔点 等小工具

tvba_core_document.py         ← 编排器:apply_settings_to_document()
tvba_core_body.py             ← 正文格式(对应 VBA RefreshContentFormat 的 BodyText 分支)
tvba_core_title.py            ← 5 级标题识别 + 格式
tvba_core_numbering.py        ← 多级列表渲染编号读取(COM 桥接)
tvba_core_toc.py              ← 目录识别 + 样式(对应 RefreshDirectoryFormat 等)
tvba_core_table.py            ← 表格 + 表格题注
tvba_core_figure.py           ← 图片题注
tvba_core_normalize.py        ← ASCII 字体统一 / ApplyBrackets / AddPeriodIfNeeded
tvba_core_oox.py              ← OOXML 底层 helpers

tvba_tests.py                 ← 测试入口(对应 wfp_tests.py)
tests/
  fixtures/                   ← 测试 .docx 样本
    build_*.py                ← 程序化构造 fixture 的脚本
    golden/                   ← VBA 跑出来的"金标准" .docx
  test_*.py                   ← 单测 + 集成 + e2e

requirements.txt
pyproject.toml
README.md
LICENSE                       ← MIT(沿用 WFP)
```

**模块对应注释约定**:每个 `tvba_core_*.py` 文件首部用注释标明对应 VBA `FormatModule.bas` 的行段,实现行级可追溯:

```python
"""tvba_core_title.py
对应 VBA FormatModule.bas:
  - AutoDetectAndFormatNumericTitles (line 783-826)
  - IdentifyContentTitleLevel (line 840-893)
  - IdentifyContentTitleLevelFromNumber (line 901-949)
  - ApplyContentTitleStyle (line 951+)
  - NormalizeNumberString (line 829-837)
"""
```

---

## 4. 数据模型

VBA `Typemodule.bas` 的扁平 Type 在 Python 改成嵌套 frozen dataclass(避免局部修改污染默认值,提升可测性)。

```python
# tvba_settings.py
from dataclasses import dataclass, field

@dataclass(frozen=True)
class TitleLevelSettings:
    """单一标题级别(对应 VBA TitleAlignment(i) / Font(i) / ... 的一项)"""
    alignment: str = "左对齐"           # 左对齐/居中/右对齐/两端对齐
    font: str = "宋体"                   # NameFarEast(中文字体)
    size: str = "小四"                   # 中文标签或纯数字字符串
    bold: bool = False
    before_lines: float = 0.5            # 段前(行)— 走 OOXML w:beforeLines
    after_lines: float = 0.5             # 段后(行)
    line_spacing: float = 1.5            # 行距倍数

@dataclass(frozen=True)
class BodySettings:
    font: str = "宋体"
    size: str = "小四"
    spacing: float = 1.5
    before_lines: float = 0.0
    after_lines: float = 0.0
    alignment: str = "两端对齐"
    left_indent_cm: float = 0.0
    right_indent_cm: float = 0.0
    special_indent: str = "首行缩进"     # 无/首行缩进/悬挂缩进
    special_indent_cm: float = 0.74      # ≈ 2 字符

@dataclass(frozen=True)
class TableSettings:
    # 表格题注(在表格上方那行 "表 1.1-1 …")
    title_font: str = "黑体"
    title_size: str = "小四"
    title_bold: bool = True
    title_spacing: float = 1.5
    title_left_indent_cm: float = 0.0
    title_right_indent_cm: float = 0.0
    title_special_indent: str = "无"
    title_special_indent_cm: float = 0.0
    # 表格本身
    body_font: str = "宋体"
    body_size: str = "五号"
    line_width_pt: float = 0.5
    row_height_cm: float = 0.7
    spacing: float = 1.0
    auto_fit_window: bool = True

@dataclass(frozen=True)
class FigureSettings:
    title_font: str = "黑体"
    title_size: str = "小四"
    title_bold: bool = True
    title_spacing: float = 1.5
    title_left_indent_cm: float = 0.0
    title_right_indent_cm: float = 0.0
    title_special_indent: str = "无"
    title_special_indent_cm: float = 0.0

@dataclass(frozen=True)
class TocLegacyFixedDefaults:
    """对应 VBA SetDefaultDirectorySettings 的写死默认值。
    UI 不暴露(VBA 同样不暴露,Mainmodule.bas 注释:"目录设置已移除"),
    但 dataclass 上保留全部字段以保证 1:1 等价。"""
    title_font: str = "宋体"
    title_size: str = "小四"
    title_bold: bool = True
    title_spacing: float = 1.5
    level1_font: str = "宋体"
    level1_size: str = "小四"
    level1_bold: bool = True
    level2_font: str = "宋体"
    level2_size: str = "小四"
    level2_indent_chars: int = 2
    level3_font: str = "宋体"
    level3_size: str = "小四"
    level3_indent_chars: int = 4

@dataclass
class FormatSettings:
    """完整格式设置 — 对应 VBA Typemodule.bas FormatSettings"""
    body: BodySettings = field(default_factory=BodySettings)
    titles: tuple[TitleLevelSettings, ...] = field(
        default_factory=lambda: tuple(TitleLevelSettings() for _ in range(5))
    )
    table: TableSettings = field(default_factory=TableSettings)
    figure: FigureSettings = field(default_factory=FigureSettings)
    toc: TocLegacyFixedDefaults = field(default_factory=TocLegacyFixedDefaults)

    auto_detect_numeric_titles: bool = True
    auto_detect_include_list_paragraphs: bool = True
    remember_settings: bool = True
```

JSON 形态:`dataclasses.asdict()` 直接序列化,层级与 dataclass 一致。

---

## 5. 核心模块

### 5.1 数据流(端到端)

```
GUI/Controller 触发 apply
    ↓
load FormatSettings(JSON / 默认值)
    ↓
[tvba_core_document.apply_settings_to_document(docx_path, settings)]
    1. python_docx.Document(docx_path) → doc
    2. body.apply_normal_style(doc, settings.body)        ← 设置 doc.styles['Normal']
    3. (可选) title.auto_detect(doc, settings, list_resolver)
                          ↓
                          numbering.resolve_list_text(para)  ← COM 或 Docx 兜底
                          ↓
                          失败 → title.identify_numeric_title_level(text)
                          ↓
                          命中 1~5 级 → 设 outlineLvl + apply_title_style
    4. for para in doc.paragraphs:
         ├─ skip if toc.is_toc_paragraph(para)
         ├─ outline_level == 0 → body.apply_paragraph(para, settings.body)
         ├─ outline_level 1..5 → title.apply_title_style(para, level, settings.titles[level-1], settings.body)
         └─ TOC → toc.apply_style(para, ...)
    5. table.refresh_all(doc, settings.table)
    6. figure.refresh_all(doc, settings.figure)
    7. normalize.unify_ascii_font(doc, "Times New Roman")
    8. doc.save(out_path)  ← 默认原地保存,等同 VBA
```

### 5.2 编排器 `tvba_core_document.py`

```python
def apply_settings_to_document(
    docx_path: Path,
    settings: FormatSettings,
    *,
    list_resolver: ListResolver | None = None,    # 不传则 numbering.auto_select()
    output_path: Path | None = None,              # 不传则原地保存
    progress_cb: Callable[[str, float], None] | None = None,
) -> Path: ...
```

不掺业务逻辑,纯调度。每一步前后调 `progress_cb`。

### 5.3 正文 `tvba_core_body.py`

VBA 对应:`RefreshContentFormat` 的 BodyText 分支 + `ApplySettingsToDocument` 里的 `wdStyleNormal` 设置。

```python
def apply_normal_style(doc, body: BodySettings) -> None
def apply_paragraph(para, body: BodySettings) -> None
```

实现要点:
- 字体走 `oox.set_far_east_font(run, name)` + `run.font.name = "Times New Roman"`
- 段前段后**优先**用 OOXML `w:beforeLines`(对应 VBA `LineUnitBefore`),只有当 LineUnit 不可用时才**降级**到 `w:before`(`SpaceBefore = lines × 行距 × 字号 × 12pt`,与 VBA 兜底逻辑一致)
- 缩进走 `oox.apply_indent_chars`(支持字符数缩进 + 厘米缩进 + 特殊缩进的"无/首行/悬挂"三选一)

### 5.4 标题 `tvba_core_title.py`(最关键模块)

VBA 对应:`AutoDetectAndFormatNumericTitles` / `IdentifyContentTitleLevel` / `IdentifyContentTitleLevelFromNumber` / `IsMultiLevelListParagraph` / `ApplyContentTitleStyle` / `NormalizeNumberString` / `ApplyBrackets` / `AddPeriodIfNeeded` / `SyncNumberFontWithBody`。

```python
def identify_numeric_title_level(text: str) -> int       # 文本 → 1~5 / 0
def identify_level_from_number(num_str: str) -> int      # 编号字符串 → 级别
def normalize_number_string(s: str) -> str               # 全角点号 / 多余尾点 / 空白清理
def auto_detect_and_format(doc, settings, list_resolver) -> None
def apply_title_style(paragraph, level, level_settings, body_settings) -> None
```

**算法保真要点**:

| VBA 行为 | 复刻方式 |
|---|---|
| `VBScript.Regexp` `^(\d+(\.\d+){0,6})[ \t]*.+$` | Python `re.match(...)` 等价 |
| `dotCount = Len(numberPart) - Len(Replace(...))` | `numberPart.count('.')` |
| 一级编号(无小数点)要求后跟空格/Tab 才认 | 同条件保留 |
| `1.0 标题` → 1 级 | `if dot_count == 1 and num.endswith('.0'): return 1` |
| 多级列表识别优先级:`ListLevelNumber`(1~5) → `IdentifyContentTitleLevelFromNumber(ListString)` → 文本兜底 | 调用顺序完全一致 |
| `OutlineLevel = wdOutlineLevel1` | `oox.set_outline_level(para, level - 1)`(OOXML 0-indexed) |
| 立即 `ApplyContentTitleStyle` | `apply_title_style` 同步执行 |

### 5.5 目录 `tvba_core_toc.py`

VBA 对应:`RefreshDirectoryFormat` / `IsTocEntryLine` / `IsTocParagraph` / `IsDirectoryTitleLine` / `IdentifyDirectoryLevel` / `ApplyTocStyleToParagraph` / `ApplyDirectoryTitleStyle` / `ApplyDirectoryStyle`。

```python
def is_toc_paragraph(para) -> bool
def is_toc_entry_line(text: str) -> bool      # 含 Tab + 末尾数字页码
def is_toc_title_line(text: str) -> bool      # 是"目录"二字行
def identify_toc_level(text: str) -> int      # 1~3 / 0
def apply_toc_title_style(para, defaults: TocLegacyFixedDefaults) -> None
def apply_toc_entry_style(doc, para, level, defaults) -> None
def refresh_toc(doc, defaults) -> None
```

实现要点:
- `is_toc_entry_line` 一字不差搬 VBA:必须含 Tab,且末尾 token 是数字(IsNumeric 等价)
- `apply_toc_entry_style` 先 `paragraph.style = doc.styles['TOC 1' / 'TOC 2' / 'TOC 3']`,再用直接格式覆盖字体字号(VBA 双重设置行为完全保留)

### 5.6 表格 `tvba_core_table.py`

VBA 对应:`RefreshTableFormat` / `SetTableTitle` / `FindTableCaptionRange` / `FindCaptionInShapes` / `IsTableCaptionLine`。

```python
def is_table_caption_line(text: str) -> bool
def find_table_caption(table, doc, max_up_paragraphs: int = 10) -> Paragraph | None
def apply_table_caption(para, settings: TableSettings) -> None
def apply_table_body(table, settings: TableSettings) -> None
def refresh_all(doc, settings: TableSettings) -> None
```

`find_table_caption` 实现两步搜索:
1. 向上找最多 10 段普通段落,命中 `is_table_caption_line` 即返回
2. 仍找不到 → 在文档级 Shape/TextFrame 里搜(对应 VBA `FindCaptionInShapes`)

`apply_table_body` 走 `oox.set_table_layout_window/content` 实现 `AutoFitBehavior`,走 `oox.set_table_borders` 实现线宽。

### 5.7 图片题注 `tvba_core_figure.py`

VBA 对应:`RefreshFigureCaptions` / `IsFigureCaptionLine`。

```python
def is_figure_caption_line(text: str) -> bool
def apply_figure_caption(para, settings: FigureSettings) -> None
def refresh_all(doc, settings: FigureSettings) -> None
```

最简单的模块。

### 5.8 多级列表桥接 `tvba_core_numbering.py`

VBA 对应:`IsMultiLevelListParagraph` / `ReportAllMultiLevelListLevels`(后者也搬过去作为诊断工具)。

```python
class ListResolver(Protocol):
    def get_list_level(self, para) -> int | None    # 1..9
    def get_list_text(self, para) -> str | None     # 渲染编号 "1.2.3"
    def diagnose(self, doc) -> list[DiagnosticEntry]

class ComListResolver:
    """通过 pywin32 启动 Word/WPS,真实读取 lf.ListLevelNumber/ListString"""
    def __enter__(self): ...
    def __exit__(self, *args): ...

class DocxListResolver:
    """纯 python-docx 兜底:读 numPr/ilvl,模拟计数。无法确认时返回 None。"""

def auto_select(prefer_com: bool = True) -> ListResolver:
    """探测 win32com.client.Dispatch('Word.Application');不可用则降级"""
```

COM 路径下 `get_list_text` 直接拿 Word 给的 ListString → 与 VBA 输出 100% 一致。这是路径 A 的核心保真点。

### 5.9 规范化 `tvba_core_normalize.py`

VBA 对应:`NormalizeAsciiFont` / `ApplyBrackets` / `AddPeriodIfNeeded` / `SyncNumberFontWithBody`。

```python
def unify_ascii_font(doc, font_name: str = "Times New Roman") -> None
def apply_brackets(para, text: str) -> None
def add_period_if_needed(para) -> None
def sync_number_font_with_body(para) -> None
```

### 5.10 OOXML helpers `tvba_core_oox.py`

所有上层模块的 lxml 封装,对外呈现"看起来像 python-docx API"的形式。

```python
# 字体
def set_far_east_font(run, font_name: str) -> None
def set_ascii_font(run, font_name: str) -> None

# 段落格式
def set_outline_level(paragraph, level_zero_indexed: int) -> None
def apply_indent_chars(
    paragraph_format,
    *, left_chars: float, right_chars: float,
    special_kind: str,         # "无" / "首行缩进" / "悬挂缩进"
    special_chars: float,
) -> None
def set_before_after_lines(paragraph_format, *, before_lines: float, after_lines: float) -> None

# 表格
def set_table_layout_window(table) -> None    # AutoFitBehavior=2
def set_table_layout_content(table) -> None   # AutoFitBehavior=1
def set_table_borders(table, *, line_width_pt: float) -> None
def set_row_height_at_least(row, height_cm: float) -> None
```

每个 helper 都是薄封装,内部用 lxml 直接操作 `<w:...>` 元素。

---

## 6. UI / Controller

### 6.1 布局(全新设计)

左侧分类树 + 右侧详情面板(类似 VS Code 设置 / macOS 系统设置)。

```
┌─────────────────────────────────────────────────────────────────────────┐
│  TransVBA-Pro — Word 格式自动刷新                              [_][□][×] │
├─────────────────────────────────────────────────────────────────────────┤
│  [📂 打开文件...]   D:\…\report.docx           [📋 预设 ▼]              │
├──────────────────┬──────────────────────────────────────────────────────┤
│  ▾ 正文          │  正文格式                                              │
│  ▾ 标题          │  ─────────────                                         │
│    ─ 1 级        │  中文字体    [宋体              ▼]                      │
│    ─ 2 级        │  字号        [小四              ▼]                      │
│  ▶ 3 级          │  行距(倍)   [1.5    ]                                  │
│    ─ 4 级        │  对齐方式    [两端对齐          ▼]                      │
│    ─ 5 级        │  ...                                                  │
│  ▾ 表格          │                                                        │
│  ▾ 图片标题      │                                                        │
│  ▾ 高级          │                                                        │
├──────────────────┴──────────────────────────────────────────────────────┤
│  ☐ 修改模式  ☑ 记忆本次设置                                               │
│                              [重置为默认]   [取消]   [应用并关闭]   [应用]  │
└─────────────────────────────────────────────────────────────────────────┘
```

UX 决策:
- **单文件模式**:每次只处理一个 .docx,等同 VBA `ActiveDocument`
- **左侧树**:VBA UserForm 的"标题"Tab 拆成 5 个独立节,5 级各点开看
- **修改模式**:对应 VBA `chkEnableEdit`("更改设置选项")— 默认只读,误触保护
- **记忆本次设置**:对应 VBA `chkRemember` — 勾选时点应用会写 JSON
- **底部按钮**:VBA 的"应用 / 确定 / 取消"三键 → 这里"应用 / 应用并关闭 / 取消"两个动作 + 重置,语义更清楚
- **预设**:加载/另存为预设(JSON 文件)— 不含 VBA Registry 导入

### 6.2 MVC 分离(TDD 友好)

```python
# Model:tvba_settings.py / tvba_persistence.py(已定义)

# Controller(完全不依赖 Tkinter)
class TvbaController:
    def __init__(self, repo: SettingsRepository, applier: DocumentApplier): ...
    @property
    def settings(self) -> FormatSettings
    @property
    def opened_file(self) -> Path | None
    def open_file(self, path: Path) -> None
    def update_setting(self, path: str, value: Any) -> ValidationResult
    def apply(self, *, save_settings: bool, progress_cb=None) -> ApplyResult
    def reset_to_defaults(self) -> None
    def load_preset(self, name: str) -> None

# View:Tkinter UI,只 import 上面这个 Controller,不直接 import core
class TvbaMainWindow(tk.Tk): ...
```

`DocumentApplier` 是抽象,生产用 `tvba_core_document.apply_settings_to_document`,测试用 mock。

---

## 7. 持久化

- **位置**:`%APPDATA%\TransVBA\tvba_config.json`(每用户,对应 VBA HKCU 注册表的语义)
- **格式**:`FormatSettings` `dataclasses.asdict()` 直接序列化(嵌套 JSON)
- **加载**:文件缺失 → 默认值;JSON 损坏 → 警告对话框 + 默认值,不阻塞 GUI 启动(对应 VBA `LoadSettingsFromRegistry` 容错)
- **保存**:用户点应用且勾选"记忆本次设置"时写入(对应 VBA `SaveSettingsToRegistry` 时机)

---

## 8. 平台与依赖

- **平台**:Windows 10/11(VBA 原本就是 Windows 路径,移植后保持)
- **运行时依赖**:
  - Python 3.11+
  - python-docx
  - pywin32(用于 numbering COM 桥接)
  - lxml(用于 OOXML helpers)
  - Tkinter(Python 标准库)
- **环境前置**:Word 或 WPS Office 已安装(用于多级列表 ListLevelNumber 读取)
- **打包**:PyInstaller,单文件 .exe(沿用 WFP 打包思路)

---

## 9. 测试策略(TDD)

### 9.1 测试金字塔

```
              E2E (~5 个)              跑完整文档 → XML 关键属性 vs VBA 金标准
        ────────────────────
       模块/集成 (~50 个)              每个 core 模块 fixture + 断言
   ────────────────────────────
  纯函数 (~150 case)                  regex / 解析 / 规范化 / helper
────────────────────────────────────
```

### 9.2 Fixture 来源(三类)

1. **VBA 注释 + 函数体反推的文本 case**:写在 Python 测试里,作为契约表
2. **程序化构造的 .docx**:`tests/fixtures/build_*.py` 脚本生成,版本可控,git 不存二进制
3. **VBA 金标准 .docx**:在 Word 里手工跑 VBA,得一份"对的"输出。**只比对关键 OOXML 属性**(字体名、字号、缩进、对齐、行距、outlineLvl 等),不比对二进制差异(`tests/fixtures/golden/*.docx`,必要时 git LFS)

### 9.3 覆盖率门槛

- core 模块 ≥ 90%
- Controller ≥ 80%
- OOXML helpers 100%(每个 helper 至少一个测试)
- View 层不做自动测试(Tkinter 在 headless CI 难跑),release 前手工 smoke

### 9.4 TDD 工作流

每个模块/特性按照 Red → Green → Refactor:

1. 从本设计与 VBA 源码反推契约,先写**失败**的测试
2. 写**最小**实现让测试绿
3. 重构,保持测试绿

实现期启用 `superpowers:test-driven-development` 技能,严格执行。

---

## 10. 风险与缓解

| 风险 | 影响 | 缓解 |
|---|---|---|
| Word 与 WPS 对同一 OOXML 解释不同(尤其表格 AutoFit / 段前段后行数) | 输出在某宿主里走样 | 每个 OOXML helper 在 Word + WPS 双宿主下手工 verify;e2e 金标准 fixture 双份 |
| 多级列表 ListLevelNumber 在没有 Office 时降级 | 个别多级列表段落识别错误级别 | 启动时探测 + GUI 状态栏明显提示;Docx Resolver 不确定时返回 None,fallback 到文本识别 |
| 用户系统缺少中文字体(宋体/黑体/方正小标宋等) | 输出字体被 Word 默认替换 | 与 VBA 同样风险,不做额外缓解,但应用前给一次性提示 |
| python-docx 处理 shape/textframe 受限 | VBA `FindCaptionInShapes` 在 docx 里能找到的 caption,Python 可能找不到 | 优先 python-docx,不行则降级到 COM 路径(同 numbering 模式) |
| 测试 fixture 维护成本 | 添加新功能时 fixture 容易过期 | 每个 fixture 配生成脚本;默认值改动用 pytest 失败暴露 |

---

## 11. 分期交付(每期 TDD 闭环)

| 期号 | 内容 | 验收 |
|---|---|---|
| **0** | 仓库脚手架 + dataclass + JSON 持久化 | `test_settings_*.py` 全绿 |
| **1** | OOXML helpers(`tvba_core_oox.py`)| `test_oox_helpers.py` 全绿,每个 helper ≥ 1 测试 |
| **2** | body 模块 | `test_body_*.py` 全绿;手工跑一份只有正文的 docx 输出对比 VBA |
| **3** | title 模块(最重模块)| `test_title_*.py` 全绿(7 个测试文件,80+ case);手工 e2e 多级标题文档 |
| **4** | toc / table / figure / normalize | 各模块测试全绿 |
| **5** | numbering(COM 桥接)| Windows + Word/WPS 环境下 `test_numbering_*.py` 全绿 |
| **6** | document 编排器 + e2e | `test_e2e.py` 跑通 5 份金标准 fixture,XML 关键属性 100% 一致 |
| **7** | Controller(GUI 业务层)| `test_controller_*.py` 全绿,UI 还没接入 |
| **8** | View(Tkinter UI)+ 手工 smoke | 主流程演示 OK |
| **9** | PyInstaller 打包 + Release | 用户机能跑 |

每期 PR 单独提交,有失败的测试就不进下一期。

---

## 12. 附录:VBA → Python 函数对照表

| VBA 函数 | Python 模块.函数 |
|---|---|
| `Mainmodule.bas:ShowFormatSettings` | `tvba.py:main` |
| `Mainmodule.bas:QuickApplyFormat` | (砍掉,GUI-only 模式用应用按钮) |
| `Typemodule.bas:FormatSettings` | `tvba_settings.py:FormatSettings` |
| `FormatModule.bas:ValidateTitleLevel5Support` | (不需要,Python dataclass 静态保证) |
| `FormatModule.bas:ApplyFormatting` | `TvbaController.apply` |
| `FormatModule.bas:GetSettingsFromForm` | View → Controller `update_setting` 调用 |
| `FormatModule.bas:LoadSettingsFromRegistry` | `tvba_persistence.py:load` |
| `FormatModule.bas:SaveSettingsToRegistry` | `tvba_persistence.py:save` |
| `FormatModule.bas:SetDefaultDirectorySettings` | `TocLegacyFixedDefaults` 默认值 |
| `FormatModule.bas:ApplySettingsToDocument` | `tvba_core_document.apply_settings_to_document` |
| `FormatModule.bas:RefreshDirectoryFormat` | `tvba_core_toc.refresh_toc` |
| `FormatModule.bas:IsDirectoryTitleLine` | `tvba_core_toc.is_toc_title_line` |
| `FormatModule.bas:ApplyDirectoryTitleStyle` | `tvba_core_toc.apply_toc_title_style` |
| `FormatModule.bas:IsTocEntryLine` | `tvba_core_toc.is_toc_entry_line` |
| `FormatModule.bas:IsTocParagraph` | `tvba_core_toc.is_toc_paragraph` |
| `FormatModule.bas:ApplyTocStyleToParagraph` | `tvba_core_toc.apply_toc_entry_style`(含 Style 设置) |
| `FormatModule.bas:IdentifyDirectoryLevel` | `tvba_core_toc.identify_toc_level` |
| `FormatModule.bas:ApplyDirectoryStyle` | `tvba_core_toc.apply_toc_entry_style`(含字体设置) |
| `FormatModule.bas:RefreshContentFormat` | `tvba_core_document.apply_settings_to_document`(里的段落循环) |
| `FormatModule.bas:IsMultiLevelListParagraph` | `tvba_core_numbering.ListResolver.get_list_level/get_list_text` |
| `FormatModule.bas:ReportAllMultiLevelListLevels` | `tvba_core_numbering.ListResolver.diagnose` |
| `FormatModule.bas:AutoDetectAndFormatNumericTitles` | `tvba_core_title.auto_detect_and_format` |
| `FormatModule.bas:NormalizeNumberString` | `tvba_core_title.normalize_number_string` |
| `FormatModule.bas:IdentifyContentTitleLevel` | `tvba_core_title.identify_numeric_title_level` |
| `FormatModule.bas:IdentifyContentTitleLevelFromNumber` | `tvba_core_title.identify_level_from_number` |
| `FormatModule.bas:ApplyContentTitleStyle` | `tvba_core_title.apply_title_style` |
| `FormatModule.bas:SafeSetPfProp` | (融入 oox helpers,不需要单独导出) |
| `FormatModule.bas:SafeGetPfPropNumber` | (同上) |
| `FormatModule.bas:ApplyParagraphIndentByChars` | `tvba_core_oox.apply_indent_chars` |
| `FormatModule.bas:SyncNumberFontWithBody` | `tvba_core_normalize.sync_number_font_with_body` |
| `FormatModule.bas:NormalizeAsciiFont` | `tvba_core_normalize.unify_ascii_font` |
| `FormatModule.bas:ApplyBrackets` | `tvba_core_normalize.apply_brackets` |
| `FormatModule.bas:AddPeriodIfNeeded` | `tvba_core_normalize.add_period_if_needed` |
| `FormatModule.bas:RefreshTableFormat` | `tvba_core_table.refresh_all` |
| `FormatModule.bas:SetTableTitle` | `tvba_core_table.apply_table_caption` |
| `FormatModule.bas:FindTableCaptionRange` | `tvba_core_table.find_table_caption`(含普通段落分支) |
| `FormatModule.bas:FindCaptionInShapes` | `tvba_core_table.find_table_caption`(含 shape/textframe 分支) |
| `FormatModule.bas:RefreshFigureCaptions` | `tvba_core_figure.refresh_all` |
| `FormatModule.bas:CleanParaText` | `tvba_utils.clean_para_text` |
| `FormatModule.bas:IsTableCaptionLine` | `tvba_core_table.is_table_caption_line` |
| `FormatModule.bas:IsFigureCaptionLine` | `tvba_core_figure.is_figure_caption_line` |
| `FormatModule.bas:ConvertSizeToPoints` | `tvba_utils.size_label_to_points` |
| `FormatModule.bas:CentimetersToPoints` | `tvba_utils.cm_to_points` |
| `UserForm1.frm:UserForm_Initialize` | `TvbaMainWindow.__init__` + `_build_layout` |
| `UserForm1.frm:CreateContentPage` | `TvbaMainWindow._build_body_panel` |
| `UserForm1.frm:CreateTitlePage` / `CreateTitleLevelGroup` | `TvbaMainWindow._build_title_panel` × 5 |
| `UserForm1.frm:CreateTablePage` | `TvbaMainWindow._build_table_panel` |
| `UserForm1.frm:CreateFigurePage` | `TvbaMainWindow._build_figure_panel` |
| `UserForm1.frm:btnApply_Click` / `btnOK_Click` / `btnCancel_Click` | `TvbaMainWindow._on_apply` / `_on_apply_close` / `_on_cancel` |
| `UserForm1.frm:LoadSettingsToForm` | View 渲染:从 Controller.settings 读 |
| `UserForm1.frm:SetEditingEnabled` | `TvbaMainWindow._set_editing_enabled` |
| `UserForm1.frm:InitializeComboBoxes` / `FillFontCombo` | `TvbaMainWindow._populate_comboboxes` |

任何 VBA 函数,都能在右列找到归属。如有遗漏,作为 spec bug 修订。

---

## 13. 下一步

1. 用户 review 本设计文档 → 反馈或确认
2. 进入 `superpowers:writing-plans`,把 §11 分期细化为可执行的实现计划
3. 实现期严格 TDD,每期 PR 独立提交

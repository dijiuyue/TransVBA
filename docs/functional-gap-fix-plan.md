# TransVBA 功能缺口修复计划

日期：2026-05-16  
目标读者：Claude 执行修复；Codex 后续 review  
目标：修复“代码、UI 或文档已经写了，但实际运行达不到预期”的功能缺口  
非目标：不追求 VBA 1:1 等价，不以旧 VBA 行为作为唯一标准

## 0. 总原则

本阶段只围绕当前产品行为修复：

1. 用户在 GUI 里能改的设置，必须真实影响输出。
2. 文档里写支持的能力，必须有代码路径和测试验证。
3. 不能实现的能力必须明确显示为 partial/unsupported，不能伪装成 supported。
4. 格式检查器不能只检查“显式写错”的 XML，也要处理样式继承或缺失属性带来的实际风险。
5. 修复时优先写行为测试，不写“为了过测试而改测试”。

## 1. 优先级总表

| 优先级 | 缺口 | 主要影响 | 主要文件 |
|---|---|---|---|
| P0 | 正文左右缩进 UI/Settings 不生效 | 用户修改后输出无变化 | `tvba_core_body.py`, `tvba_settings.py`, `tvba_gui.py`, `tests/test_body.py` |
| P0 | 表格行高检查检查不到自身输出 | 格式检查漏报 | `tvba_core_oox.py`, `tvba_core_validate.py`, `tests/test_validate.py` |
| P0 | 字体检查漏掉继承/缺失字体 | 格式检查假通过 | `tvba_core_validate.py`, `tests/test_validate.py` |
| P1 | 禁词替换不使用模板 forbidden_words | GUI 文案和实际行为不一致 | `tvba_core_body.py`, `tvba_core_document.py`, `tvba_settings.py`, `tests/test_body.py` |
| P1 | 页眉 Rev. 规范化只处理第一段 | 页眉多段时功能失效 | `tvba_core_document.py`, `tests/test_document.py` |
| P1 | 表题/图题缩进字段是假字段 | settings 有字段但格式化不使用 | `tvba_core_table.py`, `tvba_core_figure.py`, `tvba_settings.py` |
| P2 | 列表段落识别无 COM 时基本不可用 | GUI 勾选后能力有限 | `tvba_core_numbering.py`, `tvba_core_title.py`, `tvba_gui.py`, docs |
| P2 | Shape/TextFrame 表题只检测不格式化 | 表题格式化承诺不完整 | `tvba_core_table.py`, `tvba_core_document.py`, docs |
| P2 | golden_compare 说明强于实际能力 | 测试工具给虚假安全感 | `tests/fixtures/golden_compare.py` |
| P2 | 当前环境缺依赖，无法跑测试 | 无法验证实际运行 | `pyproject.toml`, `requirements.txt`, docs |

## 2. P0-1 正文左右缩进必须生效

### 当前问题

`BodySettings` 有：

- `left_indent_cm`
- `right_indent_cm`

GUI 也有对应输入框。

但 `apply_paragraph()` 里调用 `apply_indent_chars()` 时把左右缩进硬编码为 `0.0`：

```python
apply_indent_chars(
    para.paragraph_format,
    left_chars=0.0,
    right_chars=0.0,
    special_kind=body.special_indent,
    special_chars=body.special_indent_chars,
)
```

结果：用户修改正文左右缩进，输出没有变化。

### 修复方案

有两个选择，推荐方案 A。

#### 方案 A：继续使用字符单位缩进

将 GUI 文案和 settings 字段统一改成字符单位：

- `left_indent_chars`
- `right_indent_chars`

优点：和现有 `apply_indent_chars()` 一致。

缺点：需要迁移字段名。

#### 方案 B：保留 cm 字段并新增 cm 版 OOXML helper

保留现有：

- `left_indent_cm`
- `right_indent_cm`

新增：

```python
def apply_indent_cm(
    paragraph_format,
    *,
    left_cm: float,
    right_cm: float,
    special_kind: str,
    special_chars: float,
) -> None:
    ...
```

正文左右缩进使用 cm，特殊缩进仍使用字符。

推荐方案 B，因为 GUI 已经写的是 cm。

### 涉及文件

- `tvba_core_body.py`
- `tvba_core_oox.py`
- `tvba_settings.py`
- `tvba_gui.py`
- `tests/test_body.py`
- `tests/test_document.py`

### 验收测试

新增或修改测试：

1. `BodySettings(left_indent_cm=1.0, right_indent_cm=0.5)` 后，段落 XML 中：
   - `w:ind/@w:left` 约等于 1cm twips
   - `w:ind/@w:right` 约等于 0.5cm twips
2. GUI sync 后：
   - `spn_body_left=1.0`
   - `spn_body_right=0.5`
   - controller settings 真实更新。

验收标准：

- 用户在 GUI 中修改正文左右缩进，输出 docx 的正文段落缩进真实变化。

## 3. P0-2 表格行高检查要能检查自身输出

### 当前问题

格式化时：

```python
set_row_height_at_least(row, settings.row_height_cm)
```

实际写入：

```xml
w:hRule="atLeast"
```

但验证器 `_check_table_row_height()` 只在：

```python
rule == "exact"
```

时检查。

结果：工具自己设置的 `atLeast` 行高，格式检查器不会检查。

### 修复方案

选择一个产品行为：

#### 方案 A：产品要求“至少行高”

保留 `set_row_height_at_least()`，验证器接受：

- `hRule == "atLeast"`
- `hRule == "exact"`

并按规则检查：

- `atLeast`：实际值不能小于期望值太多
- `exact`：实际值应接近期望值

#### 方案 B：产品要求“固定行高”

把 helper 改成 `set_row_height_exact()`，写：

```xml
w:hRule="exact"
```

验证器保持 exact 逻辑。

推荐方案 A，避免固定行高截断内容。

### 涉及文件

- `tvba_core_oox.py`
- `tvba_core_table.py`
- `tvba_core_validate.py`
- `tests/test_validate.py`
- `tests/test_oox.py`
- `tests/test_table.py`

### 验收测试

1. 格式化后的表格行高为 `atLeast`，验证器应通过。
2. 行高低于期望值时，验证器应报 warning。
3. `exact` 行高也应按接近期望值检查。

验收标准：

- 工具自己输出的表格不会被验证器漏检。

## 4. P0-3 字体检查不能漏掉继承/缺失字体

### 当前问题

中文字体检查只在 run 有显式 `w:rFonts/@w:eastAsia` 且值错误时报错。  
ASCII 字体检查同理，只检查显式 `w:rFonts/@w:ascii`。

如果 run 没有显式字体，但实际 Word 会从样式继承错误字体，当前检查器可能误判为通过。

### 修复方案

实现有效字体解析 helper：

```python
def get_effective_run_fonts(run, paragraph, doc_styles=None) -> dict:
    return {
        "ascii": ...,
        "hAnsi": ...,
        "eastAsia": ...,
    }
```

解析优先级：

1. run direct `w:rPr/w:rFonts`
2. paragraph style run properties
3. basedOn style chain
4. document defaults
5. 若仍未知，返回 `None`

检查策略：

- 如果能解析出字体且不等于期望，报 warning。
- 如果目标文本存在，但字体无法确定，建议也报 warning，文案为“未显式设置字体，可能依赖样式继承”。

### 涉及文件

- `tvba_core_validate.py`
- `tvba_core_oox.py` 或新增 `tvba_core_effective.py`
- `tests/test_validate.py`

### 验收测试

1. run 显式错误字体：报错。
2. run 无字体，但 paragraph style 错误字体：报错。
3. run 无字体，style 正确字体：通过。
4. 完全无法确定字体：按产品决策 warning 或 ignore，但必须测试固定。

验收标准：

- “格式检查通过”不再因为 run 没显式字体而虚假通过。

## 5. P1-1 禁词替换要和模板配置一致

### 当前问题

GUI 文案写的是“同时修正内容（括号、句号、禁词替换）”。  
但 `_replace_forbidden_words()` 使用硬编码：

```python
_FORBIDDEN_MAP = {"附图": "附件", "附表": "附件"}
```

模板中的：

```python
settings.validation.forbidden_words
```

只用于检查，不用于替换。

结果：用户配置的禁词不会被“同时修正内容”替换。

### 修复方案

将内容修正从 `BodySettings.modify_content` 扩展成明确配置：

```python
@dataclass(frozen=True)
class ContentFixSettings:
    enabled: bool = False
    fix_brackets: bool = True
    add_period: bool = True
    replace_forbidden_words: bool = False
    replacements: dict[str, str] = field(default_factory=dict)
```

短期最小修复：

1. `apply_paragraph()` 不直接读硬编码 `_FORBIDDEN_MAP`。
2. 从 `FormatSettings` 或 document orchestrator 传入 replacements。
3. 如果只有 forbidden words，没有 replacement mapping，就不要自动替换，只做检查，并在 UI 文案中去掉“禁词替换”。

推荐短期方案：

- GUI 文案改为“同时修正内容（括号、句号）”
- 禁词只检查，不自动替换
- 后续如果需要自动替换，再设计 `forbidden_word_replacements`

### 涉及文件

- `tvba_core_body.py`
- `tvba_core_document.py`
- `tvba_settings.py`
- `tvba_gui.py`
- `templates/*.json`
- `tests/test_body.py`
- `tests/test_validate.py`

### 验收测试

1. `modify_content=False`：正文不被改写。
2. `modify_content=True`：括号/句号按预期修正。
3. forbidden words：只在 validate 中报 issue，不在 apply 中自动替换，除非有明确 replacement mapping。

验收标准：

- UI 文案和实际行为一致。

## 6. P1-2 页眉 Rev. 规范化处理所有页眉段落

### 当前问题

`_format_headers()` 遍历每个页眉段落时，格式化是对当前 `para` 做的。  
但规范化 `Rev.` 空格时写成了：

```python
for run in header.paragraphs[0].runs if header.paragraphs else []:
```

结果：只有第一个页眉段落会被处理。

### 修复方案

改为：

```python
for run in para.runs:
    if "Rev." in run.text:
        run.text = re.sub(r"Rev\.\s+", "Rev. ", run.text)
```

### 涉及文件

- `tvba_core_document.py`
- `tests/test_document.py`

### 验收测试

1. section header 第一个段落含 `Rev.  A`：被修成 `Rev. A`
2. section header 第二个段落含 `Rev.  B`：也被修成 `Rev. B`
3. 多 section header 都处理。

验收标准：

- 页眉任意段落中的 `Rev.` 多空格都能规范化。

## 7. P1-3 表题/图题缩进字段要么生效，要么删除

### 当前问题

配置存在：

- `title_left_indent_cm`
- `title_right_indent_cm`
- `title_special_indent`
- `title_special_indent_cm`

但格式化时硬编码清除缩进。

### 修复方案

二选一：

#### 方案 A：字段生效

表题/图题应用：

- 左缩进
- 右缩进
- 特殊缩进

注意字段单位：

- 当前字段名里有 `_cm`
- 但 `apply_indent_chars()` 接受字符单位

需要统一单位，不能直接把 cm 当 chars。

#### 方案 B：产品不支持表题/图题缩进

删除字段、删除 UI 或文档中的承诺。

推荐方案 A，设置已经存在，用户更容易理解。

### 涉及文件

- `tvba_core_table.py`
- `tvba_core_figure.py`
- `tvba_core_oox.py`
- `tvba_settings.py`
- `tvba_gui.py`
- `templates/*.json`
- `tests/test_table.py`
- `tests/test_figure.py`

### 验收测试

1. 表题设置左缩进 1cm，输出 XML 有对应 `w:ind/@w:left`。
2. 图题设置悬挂缩进，输出 XML 有对应 `w:hanging`。
3. 默认值仍保持无缩进。

验收标准：

- settings 中存在的表题/图题缩进字段真实影响输出。

## 8. P2-1 “包含列表段落”能力需要明确边界

### 当前问题

GUI 有：

- “包含列表段落”
- “使用 Word COM 读取列表级别”

但没有 Word COM 时，`DocxListResolver.get_list_text()` 返回 `None`，而标题识别不信任 fallback 的 list level。

结果：用户勾选“包含列表段落”后，如果 Word COM 不可用，自动编号标题基本不会被识别。

### 修复方案

产品决策：

1. 如果“包含列表段落”必须支持无 Word 环境：
   - 实现纯 OOXML numbering text resolver。
   - 解析 `numbering.xml`、`numId`、`abstractNum`、`lvlText`、`start`、计数器。
   - 这工作量较大。

2. 如果可以依赖 Word：
   - UI 文案改成“包含 Word 自动编号标题（需要 Word COM）”
   - 当 COM 不可用时 disable 该功能或显示明确 warning。

推荐方案 2。

### 涉及文件

- `tvba_gui.py`
- `tvba_core_numbering.py`
- `tvba_core_document.py`
- `docs/capability-matrix.md`

### 验收测试

1. COM 不可用时返回 warning。
2. warning 进入 completion dialog。
3. capability matrix 标为 partial，而不是 complete。

验收标准：

- 用户不会误以为无 Word 环境也能可靠识别自动编号标题。

## 9. P2-2 Shape/TextFrame 表题要降级明确

### 当前问题

代码能扫描 shape/text frame 中的表题，但不能格式化。  
目前会返回 warning，这是可接受的 partial 行为。

### 修复方案

本阶段不实现 shape text 格式化，只做产品边界明确：

1. capability matrix 标为 partial。
2. GUI completion warning 写清楚：
   - 检测到了文本框/形状表题
   - 当前版本不会修改
   - 需要手动处理
3. 如果将来要支持，单独做 `tvba_core_com_shapes.py`。

### 涉及文件

- `tvba_core_table.py`
- `tvba_core_document.py`
- `docs/capability-matrix.md`

### 验收测试

构造包含 shape caption 的 docx 可能较难。短期至少对 `find_table_caption(..., _shape_captions=[])` 写 XML 单测，验证 warning list 产生。

验收标准：

- 不再把 Shape/TextFrame 表题写成 supported。

## 10. P2-3 修正 golden_compare 的实际能力

### 当前问题

`golden_compare.py` 文档写会检查：

- eastAsia 字体
- szCs
- paragraph indentation

但实际没有检查这些。

### 修复方案

二选一：

1. 实现这些检查。
2. 修改 docstring，只保留实际比较项。

推荐实现，因为这些属性正是 Word 中文格式工具最关键的属性。

需要新增比较：

- `w:rFonts/@w:eastAsia`
- `w:rFonts/@w:ascii`
- `w:rFonts/@w:hAnsi`
- `w:sz`
- `w:szCs`
- `w:pPr/w:ind`
- `w:pPr/w:spacing`
- `w:pPr/w:jc`
- table row height
- table borders

### 涉及文件

- `tests/fixtures/golden_compare.py`
- `tests/test_golden_compare.py`（新增）

### 验收测试

1. 两份 docx eastAsia 不同，compare 必须报 diff。
2. 两份 docx 缩进不同，compare 必须报 diff。
3. 两份 docx szCs 不同，compare 必须报 diff。

验收标准：

- golden_compare 不再给出虚假的 matched。

## 11. P2-4 测试环境必须能跑

### 当前问题

当前环境执行：

```powershell
python -m pytest
```

失败：没有 `pytest`。

最小 smoke test 也失败：没有 `python-docx`。

### 修复方案

1. 增加 `docs/dev-setup.md`
2. 写清楚：

```powershell
python -m pip install -e ".[dev]"
python -m pytest -m "not word_com"
python -m pytest -m word_com
```

3. 如果团队习惯用 requirements：

```text
requirements-dev.txt
```

内容：

```text
-e .[dev]
```

或把 pytest 加入现有 `requirements.txt`，但更推荐保留 runtime/dev 分离。

### 涉及文件

- `pyproject.toml`
- `requirements.txt`
- `docs/dev-setup.md`
- optionally `requirements-dev.txt`

### 验收标准

- 新环境按文档能安装依赖。
- `python -m pytest -m "not word_com"` 能跑。
- 无 Word 时 `word_com` 测试能 skip，不阻塞普通测试。

## 12. 建议执行顺序

建议 Claude 按这个顺序提交，避免大改混在一起：

1. `fix: apply body left and right indentation settings`
2. `fix: validate table row height for atLeast rows`
3. `fix: resolve effective fonts during format validation`
4. `fix: align content-fix UI with actual forbidden-word behavior`
5. `fix: normalize Rev spacing across all header paragraphs`
6. `fix: apply caption indentation settings`
7. `docs: clarify COM-dependent list heading detection`
8. `test: expand golden comparison coverage`
9. `docs: add reproducible dev setup`

## 13. Codex 后续 review 清单

Claude 完成后，Codex review 时检查：

1. GUI 中每个可编辑字段是否能影响输出或明确禁用。
2. `BodySettings.left_indent_cm/right_indent_cm` 是否真实写入 OOXML。
3. `TableSettings/FigureSettings` 中所有 title 缩进字段是否真实使用。
4. 行高格式化和行高验证是否使用一致规则。
5. 字体验证是否覆盖直接格式、样式继承和未知字体。
6. 内容修正是否不会偷偷使用硬编码禁词表误改文本。
7. 页眉多段 Rev. 是否都处理。
8. COM 不可用时，列表段落功能是否明确 partial。
9. Shape/TextFrame 表题是否只声明 partial，不伪装成 supported。
10. golden_compare 是否检查它声称检查的属性。
11. `python -m pytest -m "not word_com"` 是否能在安装依赖后跑通。


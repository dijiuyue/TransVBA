# TransVBA 项目改造计划

日期：2026-05-16  
目标读者：Claude 执行实现；Codex 后续 review  
范围：只规划，不在本文件中要求立即改代码

## 0. 当前结论

这个项目现在的主要风险不是“代码少”，而是很多功能已经写成了“看起来存在”，但真实 Word/VBA 场景下达不到最初设计目标。

最初设计文档写的是：Python 桌面工具应当完整、一比一移植 VBA Word 格式刷新插件，并采用 `python-docx + pywin32 COM` 的双层策略。实际代码里，多个关键路径退化成了纯 `python-docx` 的近似处理，测试也主要验证 XML 局部属性，而不是验证真实 Word 渲染行为或 VBA golden 输出。

本计划的原则：

1. 先验真，再重构。
2. 先修会导致结果错误的核心链路，再修 UI 和体验。
3. 对“暂时做不到”的能力必须显式暴露，禁止静默降级。
4. 每个阶段都要有可运行测试和可人工验证的 Word 样本。

## 1. 项目关键路径

### 1.1 程序入口

- `tvba.py`
  - Tkinter 入口。
  - 创建 `SettingsRepository`。
  - 加载默认模板。
  - 创建 `TvbaController`。
  - 使用 `apply_settings_to_document` 作为实际格式化函数。

- `tvba_controller.py`
  - 保存当前 settings。
  - 接收 GUI 的字段更新。
  - 生成输出文件路径。
  - 调用 document applier。

- `tvba_gui.py`
  - Tkinter UI。
  - 文件打开、模板切换、应用格式、格式检查。
  - 当前仍存在 placeholder panel 机制和未完成的 preset 入口。

### 1.2 核心格式化入口

- `tvba_core_document.py`
  - `apply_settings_to_document()` 是核心编排器。
  - 当前顺序大致是：
    1. `ensure_docx()`
    2. `Document(str(docx_path))`
    3. 缓存 paragraphs/tables/styles
    4. `apply_normal_style()`
    5. `split_compound_paragraphs()`
    6. `auto_detect_and_format()`
    7. 正文/标题二次格式化
    8. `sync_numbering_with_titles()`
    9. `refresh_toc()`
    10. `refresh_tables()`
    11. `refresh_figures()`
    12. `format_cover_title()`
    13. `format_appendix()`
    14. header formatting
    15. `unify_ascii_font()`
    16. save

### 1.3 核心模块

- `tvba_core_title.py`
  - 数字标题识别。
  - 中文标题识别。
  - 列表项识别。
  - 标题格式应用。
  - compound paragraph 拆分。

- `tvba_core_numbering.py`
  - `DocxListResolver`
  - `ComListResolver`
  - `auto_select()`
  - 这是“Word 自动编号标题”是否真实可用的核心模块。

- `tvba_core_oox.py`
  - OOXML helper。
  - 字体、字号、outline、缩进、段前段后、表格边框、编号定义等底层操作。

- `tvba_core_body.py`
  - 正文格式。
  - 正文段落 normalization。

- `tvba_core_toc.py`
  - 目录识别和样式刷新。

- `tvba_core_table.py`
  - 表题识别。
  - 表格格式。
  - 文本框/Shape 表题目前只是检测意识，不是真格式化。

- `tvba_core_figure.py`
  - 图题识别和样式刷新。

- `tvba_core_validate.py`
  - 格式检查器。
  - 当前 rules 与实际 checker 不完全一致。

### 1.4 配置和模板

- `tvba_settings.py`
  - dataclass 配置模型。

- `templates/general_spec.json`
  - 通用总说明书标准。

- `templates/dapeng_internal.json`
  - 大鹏公司内部标准。

- `tvba_templates.py`
  - 读取模板 JSON。

- `tvba_persistence.py`
  - 保存用户设置到 `%APPDATA%\TransVBA\tvba_config.json`。

## 2. 关键证据和当前问题

### 2.1 测试环境本身不可复现

现象：

- 执行 `pytest` 失败：命令不存在。
- 执行 `python -m pytest` 失败：当前 Python 没有安装 pytest。

影响：

- 当前仓库不能在干净环境下一键验证。
- 后续改造没有可靠回归基线。

涉及文件：

- `pyproject.toml`
- `requirements.txt`
- `tests/`

改造目标：

- 能在项目根目录稳定运行：

```powershell
python -m pytest
```

或者明确提供：

```powershell
python -m pip install -e ".[dev]"
python -m pytest
```

验收标准：

- README 或 `docs/dev-setup.md` 中写清楚环境创建步骤。
- `python -m pytest` 可运行。
- 如果某些 COM/Word 集成测试依赖本机 Word，必须使用 marker 分离，比如：
  - `pytest -m "not word_com"`
  - `pytest -m word_com`

### 2.2 没有真实 VBA golden fixtures

现象：

- 设计文档要求 `tests/fixtures/golden/*.docx`。
- 实际仓库没有 `tests/fixtures/golden/`。
- 现有 `tests/test_e2e.py` 是现场生成极简 docx，再检查几个 XML 属性。

影响：

- 无法证明输出与 VBA 一致。
- 很容易出现“测试全绿，但 Word 里效果不对”。

涉及文件：

- `docs/superpowers/specs/2026-05-08-transvba-port-design.md`
- `tests/test_e2e.py`
- `tests/fixtures/build_full_doc.py`

改造目标：

- 新建真实验收样本目录：

```text
tests/fixtures/source/
tests/fixtures/golden/
tests/fixtures/actual/
```

建议样本：

1. `numeric_headings.docx`
   - 普通文本标题：`1 标题`、`1.1 标题`、`1.1.1 标题`
   - 验证 outline level、字体、字号、加粗、段前段后、行距。

2. `word_multilevel_list_headings.docx`
   - Word 自动多级列表标题。
   - 标题编号不是正文文本，而是 Word 渲染出来的 ListString。
   - 这是当前最容易失真的场景。

3. `toc.docx`
   - Word 自动目录。
   - 手写 tab + page number 目录。
   - 自定义 TOC 样式。

4. `table_and_figure_captions.docx`
   - 普通段落中的表题/图题。
   - 表格上方 10 段内有 caption。
   - caption 与正文、标题相邻，验证不会误判。

5. `shape_caption.docx`
   - 文本框/Shape 内表题。
   - 用于验证是否真正实现 VBA `FindCaptionInShapes`。

6. `appendix_cover_header.docx`
   - 封面标题。
   - 附件标题和附件正文。
   - 页眉 `Rev.` 文本。

验收标准：

- 每个 source 都有一份 VBA 处理后的 golden。
- Python 输出只比较关键 OOXML 属性，不做二进制全量 diff。
- 对于 Word 渲染依赖的字段，优先通过 COM 读取 Word 最终属性。

### 2.3 多级编号标题识别默认不可用

现象：

- `tvba_core_numbering.py` 中 `DocxListResolver.get_list_text()` 明确返回 `None`。
- `DocxListResolver` 注释写明无法可靠计算 rendered list text。
- `tvba_core_document.py` 只有在 `settings.prefer_com_resolver=True` 时才尝试 COM。
- 两个模板里 `prefer_com_resolver` 都是 `false`。

影响：

- 设计中最核心的“Word 自动多级列表标题识别”默认不会真正工作。
- 如果用户的标题编号是 Word 自动编号，而不是直接写在段落文本里，当前代码很可能识别失败。
- 这属于“功能看起来实现了，实际默认达不到效果”的 P0 问题。

涉及文件：

- `tvba_core_numbering.py`
- `tvba_core_document.py`
- `tvba_core_title.py`
- `templates/general_spec.json`
- `templates/dapeng_internal.json`

改造目标：

1. 明确 resolver 策略：

   - 如果启用 `auto_detect_include_list_paragraphs`，默认应使用 COM resolver。
   - 如果 Word COM 不可用，不能静默退回并宣称完成；要返回 warning 或 failure。
   - fallback 只能用于“识别 ilvl 存在”，不能用于“声称拿到了真实标题编号”。

2. 增加能力状态对象：

```python
@dataclass
class ResolverStatus:
    mode: str  # "com" | "docx_fallback" | "none"
    reliable_rendered_text: bool
    warnings: list[str]
```

3. `auto_select()` 不应只返回 resolver，还应能让调用方知道可靠性。

4. GUI 应展示：

   - 已使用 Word COM 解析自动编号。
   - 未检测到 Word COM，自动编号标题可能不会被识别。

验收标准：

- `word_multilevel_list_headings.docx` 在 COM 可用机器上通过。
- COM 不可用时，应用不会悄悄输出“看似成功但标题没处理”的文件；至少要在 ApplyResult/GUI 中给 warning。

### 2.4 COM 段落映射和 split 顺序存在结构性风险

现象：

- `apply_settings_to_document()` 先用 python-docx 打开 docx。
- 然后 `split_compound_paragraphs()` 修改内存中的 paragraph 结构。
- 之后才调用 `auto_select(..., docx_path=str(docx_path), doc=doc)`。
- `ComListResolver` 打开的是磁盘上的原始 `docx_path`。
- `ComListResolver` 用 `enumerate(doc.paragraphs)` 把 python-docx 的 paragraph id 映射到 Word COM 的 `Paragraphs(i + 1)`。

影响：

- 如果 split 后内存段落数量、顺序和磁盘原文档不同，COM 段落映射会错。
- 这会导致读取到错误段落的 ListString/ListLevelNumber。

涉及文件：

- `tvba_core_document.py`
- `tvba_core_title.py`
- `tvba_core_numbering.py`

改造方案二选一：

方案 A：预处理后保存临时 docx，再打开 COM。

流程：

1. `Document(docx_path)`
2. `split_compound_paragraphs()`
3. 保存到临时文件，例如 `%TEMP%/transvba_preprocessed_xxx.docx`
4. `ComListResolver(temp_path, doc=doc)`
5. 格式化
6. 保存最终输出
7. 删除临时文件

优点：

- python-docx 内存文档与 COM 打开的文件一致。

缺点：

- 多一次保存。
- 需要处理临时文件清理。

方案 B：COM 解析发生在 split 之前。

流程：

1. 打开原文档 COM。
2. 缓存每个原始段落的 rendered list info。
3. split 时保留来源段落 mapping。
4. 后续标题识别使用缓存。

优点：

- 少一次保存。

缺点：

- 实现复杂。
- split 产生的新段落如何继承编号信息需要定义。

建议：

- 先采用方案 A，简单可靠。

验收标准：

- compound title 样本和 Word 自动编号样本同时存在时，编号识别不串段。
- 增加测试：split 前后段落数变化，COM resolver 仍读取正确 list text。

### 2.5 文本框/Shape 表题功能是空壳

现象：

- `tvba_core_table.find_table_caption()` 中有 shape/text frame 搜索逻辑。
- 但发现 caption 后返回 `None`，因为 python-docx 不能返回可格式化 paragraph。

影响：

- VBA 的 `FindCaptionInShapes` 实际没有被等价实现。
- 用户如果表题在文本框里，当前不会格式化。
- 代码注释会让维护者误以为“已经支持 shape 检测”。

涉及文件：

- `tvba_core_table.py`
- `docs/superpowers/specs/2026-05-08-transvba-port-design.md`

改造目标：

1. 如果目标仍是 VBA 等价：
   - 增加 COM shape formatter。
   - 遍历 Word document shapes/text frames。
   - 找到表题文本。
   - 应用字体、字号、加粗、居中、行距。

2. 如果暂时不做：
   - 删除“看似支持”的分支，或者改成明确 warning。
   - validate/apply 结果中报告：检测到文本框表题，但当前版本不能修改。

验收标准：

- `shape_caption.docx` 能被真实格式化；或者明确产生 warning。
- 不允许继续静默跳过。

### 2.6 ValidationRules 与实际检查器不一致

现象：

- `ValidationRules` 里有 `check_chairman_number`。
- 模板 `general_spec.json` 里打开了 `check_chairman_number`。
- `validate_document()` 没有调用对应 checker。

影响：

- 用户以为检查开启了，实际没有检查。
- 模板配置和验证器行为不一致。

涉及文件：

- `tvba_settings.py`
- `tvba_core_validate.py`
- `templates/general_spec.json`
- `templates/dapeng_internal.json`
- `tests/test_validate.py`

改造目标：

1. 建立规则映射表：

```text
ValidationRules 字段 -> checker 函数 -> UI 是否展示 -> 模板是否使用 -> 测试文件
```

2. 对每个字段做决定：

   - 实现。
   - 移除。
   - 标记为 future，但 UI/模板不能启用。

3. `validate_document()` 里不允许有启用但无效果的规则。

验收标准：

- 每个 `ValidationRules` 字段都有测试。
- 模板中启用的规则都能触发实际 checker。
- `check_chairman_number` 要么实现，要么从模板中关闭并说明原因。

### 2.7 OOXML 写入逻辑重复且不一致

现象：

- 正文、标题、目录、图题、表题都各自写 alignment、spacing、font。
- 有些函数遍历 `para.runs`，有些函数用 `format_all_runs_in_paragraph()` 遍历嵌套 run。
- 字体、字号、加粗、行距写法分散。

影响：

- 同一设置在不同模块效果不一致。
- 域代码、hyperlink、TOC、文本框中的 run 容易漏处理。
- 后续修 bug 会扩散到多个模块。

涉及文件：

- `tvba_core_body.py`
- `tvba_core_title.py`
- `tvba_core_toc.py`
- `tvba_core_table.py`
- `tvba_core_figure.py`
- `tvba_core_oox.py`

改造目标：

新增或整理一个统一 formatting 层，例如：

```python
@dataclass(frozen=True)
class ParagraphFormatSpec:
    eastasia_font: str
    ascii_font: str = "Times New Roman"
    size_pt: float = 12.0
    bold: bool = False
    alignment: str | None = None
    before_lines: float | None = None
    after_lines: float | None = None
    line_spacing: float | None = None
    left_chars: float = 0.0
    right_chars: float = 0.0
    special_kind: str = "无"
    special_chars: float = 0.0
```

然后提供：

```python
apply_paragraph_format(para, spec)
apply_table_cell_format(cell, spec)
apply_shape_text_format(com_shape, spec)
```

验收标准：

- 正文、标题、图题、表题、目录不再各自手写完整 spacing/font 逻辑。
- 原测试全绿。
- 新 golden 测试通过。

### 2.8 apply_normal_style 不能真正设置 Normal 的中文字体

现象：

- `apply_normal_style()` 设置 `normal.font.name = "Times New Roman"`。
- 然后只遍历第一个段落的 runs 去设置 East Asian font。
- 注释写“Only need to set on style, but python-docx style font lacks eastAsia”，但实际并没有设置 Normal style 的 eastAsia。

影响：

- 如果段落没有 run，或者后续新增文本依赖 Normal style，中文字体可能不生效。
- 这类问题在 Word 里更明显，XML 局部测试不一定能发现。

涉及文件：

- `tvba_core_body.py`
- `tvba_core_oox.py`

改造目标：

- 增加 `set_style_fonts(style, ascii_font, eastasia_font, size_pt)`。
- 直接操作 style 的 `w:rPr/w:rFonts`。
- 不要依赖“第一个段落第一个 run”。

验收标准：

- Normal style XML 中有 `w:rFonts w:ascii/w:hAnsi/w:eastAsia`。
- 空文档、只有表格文档、普通正文文档都能正确继承。

### 2.9 文本替换和 normalization 有误伤风险

现象：

- `apply_paragraph()` 中默认会：
  - `apply_brackets()`
  - `add_period_if_needed()`
  - `_replace_forbidden_words()`
- `_replace_forbidden_words()` 是硬编码映射，不读取模板的 forbidden words。
- validation 有 forbidden words，但 formatting 阶段又主动替换，职责混在一起。

影响：

- 用户只是想格式化时，正文内容可能被自动改写。
- 模板检查和自动修改语义不一致。

涉及文件：

- `tvba_core_body.py`
- `tvba_core_normalize.py`
- `tvba_core_validate.py`
- `tvba_settings.py`

改造目标：

- 分离“格式刷新”和“内容修正”。
- 默认格式刷新不应改正文内容，除非模板明确要求并且 UI 可见。
- forbidden words 应只检查，还是自动替换，要产品层明确。

验收标准：

- 用户能知道哪些操作会改文本内容。
- tests 覆盖“只格式化不改内容”的默认路径。

### 2.10 Controller 输出路径需要补边界

现象：

- `TvbaController.apply()` 输出路径固定为：

```python
stem + "+格式修改后" + suffix
```

影响：

- 重复运行会覆盖同名输出。
- 如果输入是 `.doc`，`ensure_docx()` 会生成中间 `.docx`，最终 output path 和真实保存格式需要确认。
- 如果输出文件正被 Word 打开，错误提示需要可读。

涉及文件：

- `tvba_controller.py`
- `tvba_core_convert.py`
- `tvba_core_document.py`

改造目标：

- 输出路径策略集中到一个函数。
- 遇到已存在文件时：
  - 默认追加时间戳，或
  - 明确询问覆盖。
- `.doc` 输入要明确输出 `.docx`，不要仍用 `.doc` 后缀。

验收标准：

- `.docx` 输入输出 `.docx`。
- `.doc` 输入输出 `.docx`。
- 输出文件被占用时 GUI 给出中文可读提示。

### 2.11 GUI 不是当前首要，但需要收尾

现象：

- `load_preset()` 是 TODO。
- `_build_placeholder_panel()` 仍存在。
- 高级选项里 `prefer_com_resolver` 对核心行为影响巨大，但默认模板关闭。

涉及文件：

- `tvba_controller.py`
- `tvba_gui.py`

改造目标：

- P0/P1 核心能力稳定后再做 UI。
- GUI 中必须展示能力状态：
  - 当前是否使用 Word COM。
  - 是否检测到无法处理的 Shape caption。
  - 是否有自动编号标题未可靠识别。

验收标准：

- 用户点“应用”后，不只看到成功/失败，还能看到 warning。
- placeholder 不再对应任何用户可点击的真实功能。

## 3. 建议执行阶段

## Phase 0：冻结现状和建立基线

目标：让后续所有改造都有可验证基准。

任务：

1. 新建 `docs/capability-matrix.md`。
2. 运行环境修复。
3. 建立 pytest marker。
4. 建立 source/golden/actual fixture 目录。
5. 写 golden 比较 helper。

建议提交：

```text
chore: add reproducible test setup and capability matrix
test: add golden fixture structure and comparison helpers
```

验收：

- `python -m pytest -m "not word_com"` 可运行。
- 没有 Word 的机器也能跑非 COM 测试。
- Word/COM 测试可以单独跑。

## Phase 1：修自动编号标题识别

目标：解决默认最容易失真的核心功能。

任务：

1. 重写 `auto_select()` 返回 resolver + status。
2. 默认策略改为：
   - 需要 list paragraph rendered text 时优先 COM。
   - COM 不可用则 warning。
3. 改 `apply_settings_to_document()`，把 warning 传回 controller。
4. 改 GUI 显示 warning。
5. 增加 Word 自动多级列表 fixture。

建议提交：

```text
fix: make multilevel numbering use reliable COM resolver
test: add Word multilevel list heading golden coverage
```

验收：

- Word 自动编号标题能识别为 1-5 级标题。
- 普通正文列表不会误判为标题。
- COM 不可用时不会静默假成功。

## Phase 2：修 COM 映射和 split 顺序

目标：消除 python-docx 内存文档与 Word COM 磁盘文档不一致的问题。

任务：

1. 采用“split 后保存临时 docx，再打开 COM”的策略。
2. 增加临时文件管理。
3. 确保异常时也关闭 COM、清理临时文件。
4. 增加 split + 自动编号混合 fixture。

建议提交：

```text
fix: align COM resolver with preprocessed document state
test: cover compound headings with numbered list resolution
```

验收：

- split 后段落与 COM 段落不会错位。
- 异常时没有残留 Word 进程。

## Phase 3：真实处理 Shape/TextFrame 表题

目标：处理或显式降级 VBA `FindCaptionInShapes`。

任务：

1. 做技术决策：
   - 实现 COM shape formatting，或
   - 明确不支持并 warning。
2. 如果实现：
   - 新增 `tvba_core_com_shapes.py`。
   - 遍历 Word shapes/text frames。
   - 匹配 table caption。
   - 写入字体、字号、加粗、对齐、行距。
3. 如果暂不实现：
   - 删除 misleading 的 `return None` 假支持。
   - 在 apply/validate 输出 warning。

建议提交：

```text
fix: handle table captions inside Word shapes
```

或：

```text
fix: report unsupported shape captions explicitly
```

验收：

- `shape_caption.docx` 有确定结果。
- 不允许静默跳过。

## Phase 4：对齐验证器

目标：模板启用的检查必须真实执行。

任务：

1. 建 `ValidationRules` 映射表。
2. 实现或移除 `check_chairman_number`。
3. 每个 checker 增加独立单测。
4. `validate_document()` 返回 issue code，避免只靠中文 description 判断。

建议 issue model：

```python
@dataclass
class ValidationIssue:
    code: str
    severity: str
    description: str
    location: str = ""
```

建议提交：

```text
fix: align validation rules with implemented checks
test: cover every enabled validation rule
```

验收：

- 模板中打开的规则全部可触发。
- 没有 unused validation flag。

## Phase 5：统一 OOXML 格式写入

目标：减少分散实现导致的不一致。

任务：

1. 在 `tvba_core_oox.py` 或新文件中建立统一 formatter。
2. 正文、标题、目录、图题、表题逐步迁移。
3. 每迁移一个模块就跑对应测试和 golden。
4. 不要一次性大爆炸重构。

建议提交：

```text
refactor: centralize paragraph formatting helpers
refactor: migrate title formatting to shared formatter
refactor: migrate caption and toc formatting to shared formatter
```

验收：

- 格式写入行为一致。
- 原有测试和 golden 测试通过。

## Phase 6：清理内容修改职责

目标：区分格式化和文本内容修改。

任务：

1. 梳理以下函数是否应默认执行：
   - `apply_brackets()`
   - `add_period_if_needed()`
   - `_replace_forbidden_words()`
2. 增加 settings 开关。
3. GUI 展示这些开关。
4. 默认策略建议：
   - 格式刷新默认不改正文内容。
   - 内容修正需要用户明确开启。

建议提交：

```text
fix: separate content normalization from format refresh
```

验收：

- 默认 apply 不会意外替换正文词语。
- 用户开启内容修正后才执行文本改写。

## Phase 7：完善 GUI 和输出体验

目标：让用户知道软件到底做了什么。

任务：

1. ApplyResult 增加 warnings。
2. GUI 展示 warning 摘要。
3. 输出路径策略改造。
4. 预设加载/保存闭环。
5. COM 不可用提示。
6. 文件被占用提示。

建议提交：

```text
feat: surface apply warnings in GUI
feat: complete preset load and save workflow
fix: improve output path collision handling
```

验收：

- 用户不会把部分失败误认为成功。
- 重复运行不会无提示覆盖输出。

## Phase 8：打包前最终验收

目标：确认 exe 只是交付形式，不掩盖功能问题。

任务：

1. 跑完整测试：

```powershell
python -m pytest
python -m pytest -m word_com
```

2. 人工打开 actual 输出文档检查：
   - 标题
   - 多级编号
   - 目录
   - 图题
   - 表题
   - Shape caption
   - 封面
   - 附件
   - 页眉

3. 再运行 `build_exe.ps1`。

建议提交：

```text
build: prepare verified PyInstaller package
```

验收：

- exe 在一台干净 Windows + Word 环境能处理 golden source 样本。
- 结果与 Python 直接运行一致。

## 4. Claude 执行时的注意事项

1. 不要先美化 UI。
2. 不要先重命名所有文件。
3. 不要把 COM 失败吞掉。
4. 不要用 mock 代替真实 Word 行为作为最终验收。
5. 不要新增“看似支持但实际 return None”的能力分支。
6. 每个 P0/P1 修复都要附带至少一个真实 docx fixture。
7. 所有 warnings 都要能从 controller 传到 GUI。
8. 修改输出文档内容的功能必须和格式化功能分开。

## 5. Codex 后续 review 检查清单

Claude 完成后，Codex review 时重点检查：

1. `python -m pytest` 是否能在当前环境运行。
2. 是否新增真实 golden fixtures。
3. Word 自动多级编号是否默认可靠。
4. COM 不可用时是否有 warning。
5. split 后 COM 段落映射是否可靠。
6. Shape/TextFrame 表题是否真实处理或明确 warning。
7. 模板 validation rules 是否全部有对应 checker。
8. 正文格式刷新是否会意外改正文内容。
9. 输出路径是否安全。
10. GUI 是否展示部分失败/降级信息。

## 6. 推荐优先级总表

| 优先级 | 工作项 | 主要文件 | 为什么优先 |
|---|---|---|---|
| P0 | 测试环境可复现 | `pyproject.toml`, `requirements.txt`, `tests/` | 没有回归基线，后面都不稳 |
| P0 | golden fixtures | `tests/fixtures/` | 当前测试不能证明 Word/VBA 等价 |
| P0 | 自动编号标题 COM resolver | `tvba_core_numbering.py`, `tvba_core_document.py`, `tvba_core_title.py` | 核心功能默认不可靠 |
| P0 | COM 与 split 顺序 | `tvba_core_document.py`, `tvba_core_title.py`, `tvba_core_numbering.py` | 会造成段落错位和错误识别 |
| P1 | Shape/TextFrame 表题 | `tvba_core_table.py` | 设计承诺但实际没有格式化 |
| P1 | validation rules 对齐 | `tvba_settings.py`, `tvba_core_validate.py`, `templates/*.json` | 用户以为检查了，实际没检查 |
| P1 | 统一 OOXML formatter | `tvba_core_oox.py`, core modules | 减少格式不一致 |
| P2 | 内容修改职责分离 | `tvba_core_body.py`, `tvba_core_normalize.py` | 防止格式化时误改正文 |
| P2 | GUI warning 和 preset | `tvba_gui.py`, `tvba_controller.py` | 核心稳定后再完善体验 |
| P3 | 打包 exe | `build_exe.ps1` | 最后交付，不应提前 |


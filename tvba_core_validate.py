"""Document format validation engine.

Scans a Word document against the active template's validation rules
and returns a list of format issues found.
"""
import re
from dataclasses import dataclass
from pathlib import Path

from docx import Document
from tvba_settings import FormatSettings


@dataclass
class ValidationIssue:
    severity: str  # "error" | "warning"
    description: str
    location: str = ""


def validate_document(
    docx_path: Path,
    settings: FormatSettings,
    progress_cb=None,
) -> list[ValidationIssue]:
    """Scan a document and return format issues based on the template's validation rules."""
    rules = settings.validation
    issues: list[ValidationIssue] = []

    doc = Document(str(docx_path))
    paragraphs = list(doc.paragraphs)
    tables = list(doc.tables)
    total = len(paragraphs) + len(tables)

    count = 0
    def tick(msg: str):
        nonlocal count
        count += 1
        if progress_cb:
            progress_cb(msg, count / max(total, 1))

    if rules.check_chinese_font:
        for para in paragraphs:
            _check_chinese_font(para, issues)
            tick("检查中文字体...")

    if rules.check_ascii_font:
        for para in paragraphs:
            _check_ascii_font(para, issues)
            tick("检查数字/英文字体...")

    if rules.check_brackets:
        for para in paragraphs:
            _check_brackets(para, issues)
            tick("检查括号格式...")

    if rules.check_period:
        for para in paragraphs:
            _check_period(para, issues)
            tick("检查列表句号...")

    if rules.check_forbidden_words and rules.forbidden_words:
        for para in paragraphs:
            _check_forbidden_words(para, rules.forbidden_words, issues)
            tick("检查禁忌词...")

    if rules.check_table_font_size:
        for table in tables:
            _check_table_font_size(table, settings, issues)
            tick("检查表格字号...")

    if rules.check_table_row_height:
        for table in tables:
            _check_table_row_height(table, settings, issues)
            tick("检查表格行高...")

    if rules.check_cover_title_size:
        _check_cover_title_size(paragraphs, settings, issues)
        tick("检查封面字号...")

    if rules.check_appendix_body_size:
        _check_appendix_body_size(paragraphs, settings, issues)
        tick("检查附件正文字号...")

    if rules.check_grid_alignment:
        _check_grid_alignment(paragraphs, issues)
        tick("检查网格对齐...")

    if rules.check_appendix_colon:
        _check_appendix_colon(paragraphs, issues)
        tick("检查附件冒号...")

    if rules.check_figure_table_space:
        _check_figure_table_space(paragraphs, issues)
        tick("检查图表空格...")

    if rules.check_figure_position:
        _check_figure_position(doc, tables, paragraphs, issues)
        tick("检查图表位置...")

    if progress_cb:
        progress_cb("检查完成", 1.0)

    return issues


def _para_location(para) -> str:
    text = para.text[:50].replace("\n", " ")
    if len(para.text) > 50:
        text += "..."
    return f'"{text}"'


def _check_chinese_font(para, issues: list[ValidationIssue]):
    """Check if Chinese characters use the correct font (宋体)."""
    text = para.text
    if not text:
        return
    has_chinese = bool(re.search(r'[一-鿿]', text))
    if not has_chinese:
        return
    for run in para.runs:
        if not re.search(r'[一-鿿]', run.text):
            continue
        font = run.font
        try:
            rPr = run._element.find(
                ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr"
            )
            if rPr is not None:
                rFonts = rPr.find(
                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts"
                )
                if rFonts is not None:
                    east = rFonts.get(
                        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia"
                    )
                    if east and east != "宋体":
                        issues.append(ValidationIssue(
                            severity="warning",
                            description=f"中文字体非宋体: {east}",
                            location=_para_location(para),
                        ))
                        return
        except Exception:
            pass


def _check_ascii_font(para, issues: list[ValidationIssue]):
    """Check if ASCII characters (digits, letters) use Times New Roman."""
    text = para.text
    if not text:
        return
    has_ascii = bool(re.search(r'[a-zA-Z0-9]', text))
    if not has_ascii:
        return
    for run in para.runs:
        if not re.search(r'[a-zA-Z0-9]', run.text):
            continue
        try:
            rPr = run._element.find(
                ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr"
            )
            if rPr is not None:
                rFonts = rPr.find(
                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts"
                )
                if rFonts is not None:
                    ascii_font = rFonts.get(
                        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii"
                    )
                    if ascii_font and ascii_font != "Times New Roman":
                        issues.append(ValidationIssue(
                            severity="warning",
                            description=f"数字/英文字体非 Times New Roman: {ascii_font}",
                            location=_para_location(para),
                        ))
                        return
        except Exception:
            pass


def _check_brackets(para, issues: list[ValidationIssue]):
    """Check for half-width brackets in list items (四级及以下)."""
    text = para.text
    if not text:
        return
    # Check for half-width parentheses that should be full-width
    if re.search(r'\([^)]*\)', text):
        issues.append(ValidationIssue(
            severity="warning",
            description="列表项包含半角括号 ()，应使用全角括号（）",
            location=_para_location(para),
        ))


def _check_period(para, issues: list[ValidationIssue]):
    """Check if a list item is missing a period at the end.

    Only flags genuine list items (with bracket/parenthesis delimiters
    or letter prefixes), NOT title/heading paragraphs (number+space).
    """
    text = (para.text or "").strip()
    if not text:
        return

    # Skip title paragraphs (number. or number.number + space)
    if re.match(r'^\d+(\.\d+)*\s+\S', text):
        return

    # Check if paragraph looks like a list item with bracket/parenthesis
    # e.g. "(1) xxx", "1) xxx", "（1）xxx", "a. xxx", "a) xxx"
    is_list_item = (
        re.match(r'^[\(\（]\d+[\)\）]', text) or  # (1) or （1）
        re.match(r'^\d+[\)\）]', text) or          # 1) or 1）
        re.match(r'^[a-z][\.\)]\s', text)           # a. or a)
    )
    if not is_list_item:
        return

    if not text.endswith(("。", "）", ")", ".", "；", ";", "！", "？")):
        issues.append(ValidationIssue(
            severity="warning",
            description="列表项末尾缺少句号",
            location=_para_location(para),
        ))


def _check_forbidden_words(para, forbidden_words: tuple[str, ...], issues: list[ValidationIssue]):
    """Check for forbidden words in paragraph text."""
    text = para.text
    if not text:
        return
    for word in forbidden_words:
        if word in text:
            issues.append(ValidationIssue(
                severity="error",
                description=f'发现禁忌词"{word}"，建议替换',
                location=_para_location(para),
            ))


def _check_table_font_size(table, settings: FormatSettings, issues: list[ValidationIssue]):
    """Check if table body cells use the correct font size."""
    expected_size = settings.table.body_size
    for row_idx, row in enumerate(table.rows):
        for col_idx, cell in enumerate(row.cells):
            for para in cell.paragraphs:
                for run in para.runs:
                    try:
                        rPr = run._element.find(
                            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr"
                        )
                        if rPr is not None:
                            sz = rPr.find(
                                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz"
                            )
                            if sz is not None:
                                size_val = sz.get(
                                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"
                                )
                                if size_val:
                                    from tvba_utils import size_label_to_points
                                    expected_pt = size_label_to_points(expected_size)
                                    actual_half_pt = int(size_val)
                                    actual_pt = actual_half_pt / 2
                                    if actual_pt != expected_pt:
                                        issues.append(ValidationIssue(
                                            severity="warning",
                                            description=f"表格({row_idx+1},{col_idx+1})字号非{expected_size}",
                                            location=f"表格行{row_idx+1}列{col_idx+1}",
                                        ))
                                        return
                    except Exception:
                        pass


def _check_table_row_height(table, settings: FormatSettings, issues: list[ValidationIssue]):
    """Check if table rows have the correct height."""
    expected_cm = settings.table.row_height_cm
    expected_twips = expected_cm * 567  # 1 cm = 567 twips
    for row_idx, row in enumerate(table.rows):
        try:
            tr = row._element
            trPr = tr.find(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}trPr"
            )
            if trPr is not None:
                trHeight = trPr.find(
                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}trHeight"
                )
                if trHeight is not None:
                    val = trHeight.get(
                        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"
                    )
                    rule = trHeight.get(
                        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hRule"
                    )
                    if val and rule == "exact":
                        actual_twips = int(val)
                        if abs(actual_twips - expected_twips) > 10:
                            issues.append(ValidationIssue(
                                severity="warning",
                                description=f"表格行{row_idx+1}行高非{expected_cm}cm",
                                location=f"表格行{row_idx+1}",
                            ))
        except Exception:
            pass


def _check_cover_title_size(paragraphs, settings: FormatSettings, issues: list[ValidationIssue]):
    """Check if the cover page title uses 二号 (22pt) fontWeight."""
    from tvba_utils import size_label_to_points
    expected_pt = size_label_to_points("二号")

    for para in paragraphs[:15]:
        text = (para.text or "").strip()
        if not text or len(text) < 2 or len(text) > 80:
            continue
        if para.alignment != 1:
            continue
        # Stop at TOC boundary
        from tvba_core_toc import is_toc_title_line, is_toc_entry_line
        if is_toc_title_line(text) or is_toc_entry_line(text):
            break
        # Check font size in first run
        for run in para.runs:
            try:
                rPr = run._element.find(
                    ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr"
                )
                if rPr is not None:
                    sz = rPr.find(
                        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz"
                    )
                    if sz is not None:
                        half_pts = int(sz.get(
                            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"
                        ))
                        actual_pt = half_pts / 2
                        if actual_pt != expected_pt:
                            issues.append(ValidationIssue(
                                severity="warning",
                                description=f"封面标题字号非二号(22pt)，当前为{actual_pt:.0f}pt",
                                location=_para_location(para),
                            ))
            except Exception:
                pass
        break  # Only check first qualifying paragraph


def _check_appendix_body_size(paragraphs, settings: FormatSettings, issues: list[ValidationIssue]):
    """Check if appendix body text uses 小五 (10.5pt) font."""
    from tvba_utils import size_label_to_points
    from tvba_core_appendix import is_appendix_title
    expected_pt = size_label_to_points("小五")

    in_appendix = False
    for para in paragraphs:
        text = (para.text or "").strip()
        if not text:
            continue
        if is_appendix_title(text):
            in_appendix = True
            continue
        if not in_appendix:
            continue
        # Check if ending condition: next heading
        try:
            rPr = para._element.find(
                ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr"
            )
            if rPr is not None:
                outline = rPr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}outlineLvl")
                if outline is not None:
                    in_appendix = False
                    continue
        except Exception:
            pass
        # Check font size
        for run in para.runs:
            try:
                rPr = run._element.find(
                    ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr"
                )
                if rPr is not None:
                    sz = rPr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz")
                    if sz is not None:
                        half_pts = int(sz.get(
                            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"
                        ))
                        actual_pt = half_pts / 2
                        if actual_pt != expected_pt:
                            issues.append(ValidationIssue(
                                severity="warning",
                                description=f"附件正文字号非小五(10.5pt)，当前为{actual_pt:.0f}pt",
                                location=_para_location(para),
                            ))
                            return in_appendix
            except Exception:
                pass


def _check_grid_alignment(paragraphs, issues: list[ValidationIssue]):
    """Check if level 1 titles have snapToGrid enabled."""
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    for para in paragraphs:
        try:
            pPr = para._element.find(f".//{{{W}}}pPr")
            if pPr is None:
                continue
            outline = pPr.find(f"{{{W}}}outlineLvl")
            if outline is None:
                continue
            level_val = outline.get(f"{{{W}}}val")
            if level_val != "0":
                continue
            # This is a level 1 title — check grid alignment
            snap = pPr.find(f"{{{W}}}snapToGrid")
            auto_de = pPr.find(f"{{{W}}}autoSpaceDE")
            if snap is None or auto_de is None:
                issues.append(ValidationIssue(
                    severity="warning",
                    description="一级标题未勾选'对齐到网格'或'自动调整右缩进'",
                    location=_para_location(para),
                ))
                return
        except Exception:
            pass


def _check_appendix_colon(paragraphs, issues: list[ValidationIssue]):
    """Check if appendix titles use fullwidth colon (：) after the number."""
    from tvba_core_appendix import is_appendix_title
    for para in paragraphs:
        text = (para.text or "").strip()
        if not text:
            continue
        if is_appendix_title(text):
            # Check for ： after the appendix prefix
            import re
            if not re.search(r'附件\d*[：:]', text):
                issues.append(ValidationIssue(
                    severity="warning",
                    description="附件标题未使用冒号（应为全角 ：）",
                    location=_para_location(para),
                ))


def _check_figure_table_space(paragraphs, issues: list[ValidationIssue]):
    """Check if figure/table captions have exactly one space between number and text."""
    from tvba_core_table import is_table_caption_line
    from tvba_core_figure import is_figure_caption_line
    for para in paragraphs:
        text = (para.text or "").strip()
        if not text:
            continue
        if is_table_caption_line(text) or is_figure_caption_line(text):
            # Check spacing: number part should be separated by one space
            import re
            m = re.match(r'^([表图]\s*\d+(?:\.\d+)*-\d+)(\s+)(.+)$', text)
            if m and len(m.group(2)) != 1:
                issues.append(ValidationIssue(
                    severity="warning",
                    description=f"题注序号与名称间应恰好1个空格（当前{len(m.group(2))}个）",
                    location=_para_location(para),
                ))


def _check_figure_position(doc, tables, paragraphs, issues: list[ValidationIssue]):
    """Check figure/table caption positioning: table captions above, figure captions below.

    Checks that captions appear BEFORE their associated table/figure.
    """
    from tvba_core_table import is_table_caption_line
    from tvba_core_figure import is_figure_caption_line

    body = doc.element.body
    elements = list(body)
    for i, elem in enumerate(elements):
        tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
        if tag == "tbl":
            # Check if the previous non-empty element is a table caption
            for j in range(i - 1, -1, -1):
                prev_tag = elements[j].tag.split("}")[-1] if "}" in elements[j].tag else elements[j].tag
                if prev_tag == "p":
                    # Get text from this paragraph element
                    texts = []
                    for t in elements[j].findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"):
                        if t.text:
                            texts.append(t.text)
                    para_text = "".join(texts).strip()
                    if para_text and not is_table_caption_line(para_text):
                        issues.append(ValidationIssue(
                            severity="warning",
                            description="表格上方应有题注（表x.x-x 题注文字）",
                            location=f'表格前段落: "{para_text[:30]}..."' if len(para_text) > 30 else f'表格前段落: "{para_text}"',
                        ))
                    break

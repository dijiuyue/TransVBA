"""Table + table caption formatting.

Corresponds to VBA FormatModule.bas:
  - RefreshTableFormat
  - SetTableTitle
  - FindTableCaptionRange
  - FindCaptionInShapes
  - IsTableCaptionLine
"""
from tvba_core_oox import (
    set_table_layout,
    set_table_borders,
    set_table_alignment,
    set_row_height_at_least,
    apply_indent_cm,
    apply_paragraph_spacing,
    set_paragraph_alignment,
    format_all_runs_in_paragraph,
)
import re
from lxml import etree

from tvba_utils import clean_para_text, size_label_to_points, cm_to_points

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W_VAL = f"{{{W}}}val"

# Caption number category:
#   表1 标题
#   表1-1 标题
#   表1.8.12-2 标题
# Also accepts full-width digits/dots/dashes copied from Word/WPS.
_CAPTION_NUMBER_PATTERN = r"[0-9０-９]+(?:[.．][0-9０-９]+)*(?:[-－–—][0-9０-９]+)?"
_TABLE_CAPTION_RE = re.compile(
    rf"^(?:表|table)\s*{_CAPTION_NUMBER_PATTERN}\s+.+$",
    re.IGNORECASE,
)


def is_table_caption_line(text: str) -> bool:
    """Check if text is a table caption."""
    text = clean_para_text(text)
    return bool(_TABLE_CAPTION_RE.match(text))


def is_table_caption_paragraph(para, doc=None) -> bool:
    """Check whether a paragraph is a table caption.

    Word captions can store the visible "表1.8.13-1" prefix as an automatic
    numbering label rather than paragraph text. In that case paragraph.text is
    only the caption title, while w:numPr points to a numbering level whose
    w:lvlText starts with "表".
    """
    text = clean_para_text(para.text)
    if is_table_caption_line(text):
        return True
    if not text:
        return False
    if _style_is_table_caption(para, doc):
        return True
    return _numbering_label_starts_with_table(para, doc)


def _style_is_table_caption(para, doc=None) -> bool:
    """Return True for paragraph styles whose Word style name starts with 表."""
    style_id = _paragraph_style_id(para)
    if not style_id:
        return False
    style = _style_element_by_id(doc, style_id)
    if style is None:
        return False
    name = style.find(f"{{{W}}}name")
    style_name = name.get(_W_VAL, "") if name is not None else ""
    return bool(re.match(rf"^表\s*{_CAPTION_NUMBER_PATTERN}", style_name))


def _numbering_label_starts_with_table(para, doc=None) -> bool:
    """Return True if paragraph numbering format starts with 表."""
    ilvl, num_id = _effective_num_pr(para, doc)
    if ilvl is None or num_id is None:
        return False
    lvl_text = _numbering_level_text(doc, num_id, ilvl)
    return bool(lvl_text and lvl_text.lstrip().startswith("表"))


def _paragraph_style_id(para) -> str | None:
    pPr = para._element.find(f"{{{W}}}pPr")
    if pPr is None:
        return None
    pStyle = pPr.find(f"{{{W}}}pStyle")
    return None if pStyle is None else pStyle.get(_W_VAL)


def _effective_num_pr(para, doc=None) -> tuple[int | None, str | None]:
    """Read direct numbering first, then inherit numbering from paragraph style."""
    pPr = para._element.find(f"{{{W}}}pPr")
    if pPr is not None:
        num_pr = pPr.find(f"{{{W}}}numPr")
        ilvl, num_id = _num_pr_values(num_pr)
        if num_id is not None:
            return ilvl, num_id

    style_id = _paragraph_style_id(para)
    seen: set[str] = set()
    while style_id and style_id not in seen:
        seen.add(style_id)
        style = _style_element_by_id(doc, style_id)
        if style is None:
            break
        num_pr = style.find(f"{{{W}}}pPr/{{{W}}}numPr")
        ilvl, num_id = _num_pr_values(num_pr)
        if num_id is not None:
            if ilvl is None:
                outline = style.find(f"{{{W}}}pPr/{{{W}}}outlineLvl")
                ilvl = _int_or_none(outline.get(_W_VAL)) if outline is not None else 0
            return ilvl, num_id
        based_on = style.find(f"{{{W}}}basedOn")
        style_id = based_on.get(_W_VAL) if based_on is not None else None
    return None, None


def _num_pr_values(num_pr) -> tuple[int | None, str | None]:
    if num_pr is None:
        return None, None
    ilvl_elem = num_pr.find(f"{{{W}}}ilvl")
    num_id_elem = num_pr.find(f"{{{W}}}numId")
    ilvl = _int_or_none(ilvl_elem.get(_W_VAL)) if ilvl_elem is not None else None
    num_id = num_id_elem.get(_W_VAL) if num_id_elem is not None else None
    return ilvl, num_id


def _style_element_by_id(doc, style_id: str):
    styles_element = _styles_element(doc)
    if styles_element is None:
        return None
    return styles_element.find(f".//{{{W}}}style[@{{{W}}}styleId='{style_id}']")


def _styles_element(doc):
    if doc is None:
        return None
    try:
        return doc.styles.element
    except Exception:
        return None


def _numbering_level_text(doc, num_id: str, ilvl: int | None) -> str | None:
    lvl = _numbering_level(doc, num_id, ilvl)
    if lvl is None:
        return None
    lvl_text = lvl.find(f"{{{W}}}lvlText")
    return None if lvl_text is None else lvl_text.get(_W_VAL)


def _numbering_level(doc, num_id: str, ilvl: int | None):
    numbering = _numbering_element(doc)
    if numbering is None or ilvl is None:
        return None

    num = numbering.find(f".//{{{W}}}num[@{{{W}}}numId='{num_id}']")
    if num is None:
        return None

    override = num.find(f"{{{W}}}lvlOverride[@{{{W}}}ilvl='{ilvl}']")
    if override is not None:
        lvl = override.find(f"{{{W}}}lvl")
        if lvl is not None:
            return lvl

    abstract_id = num.find(f"{{{W}}}abstractNumId")
    if abstract_id is None:
        return None
    abstract = numbering.find(f".//{{{W}}}abstractNum[@{{{W}}}abstractNumId='{abstract_id.get(_W_VAL)}']")
    if abstract is None:
        return None
    return abstract.find(f"{{{W}}}lvl[@{{{W}}}ilvl='{ilvl}']")


def _numbering_element(doc):
    if doc is None:
        return None
    try:
        return doc.part.numbering_part._element
    except Exception:
        return None


def _int_or_none(value) -> int | None:
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def find_table_caption(table, doc, max_up_paragraphs: int = 10, *, _tables=None, _paragraphs=None):
    """Find the caption paragraph preceding a table."""
    # Find table index in document
    tables = _tables if _tables is not None else doc.tables
    table_index = None
    for i, t in enumerate(tables):
        if t._element is table._element:
            table_index = i
            break

    if table_index is None:
        return None

    # Count paragraphs before this table
    paragraphs_before = 0
    for element in doc.element.body:
        if element is table._element:
            break
        if element.tag.endswith("}p"):
            paragraphs_before += 1

    # Search backwards up to max_up_paragraphs
    paragraphs = _paragraphs if _paragraphs is not None else doc.paragraphs
    for i in range(1, max_up_paragraphs + 1):
        idx = paragraphs_before - i
        if idx < 0:
            break
        para = paragraphs[idx]
        if is_table_caption_paragraph(para, doc):
            return para

    return None


def _remove_outline_level(para) -> None:
    """Remove any outlineLvl element so captions are not treated as headings."""
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    pPr = para._element.find(f"{{{W}}}pPr")
    if pPr is not None:
        outline = pPr.find(f"{{{W}}}outlineLvl")
        if outline is not None:
            pPr.remove(outline)


def _set_bool_property(rPr, tag: str, value: bool) -> None:
    for old in list(rPr.findall(f"{{{W}}}{tag}")):
        rPr.remove(old)
    elem = etree.SubElement(rPr, f"{{{W}}}{tag}")
    if not value:
        elem.set(_W_VAL, "0")


def _format_numbering_level_for_caption(para, doc, settings) -> None:
    """Format Word-generated caption labels stored in numbering.xml."""
    ilvl, num_id = _effective_num_pr(para, doc)
    lvl = _numbering_level(doc, num_id, ilvl) if num_id is not None else None
    if lvl is None:
        return

    rPr = lvl.find(f"{{{W}}}rPr")
    if rPr is None:
        rPr = etree.SubElement(lvl, f"{{{W}}}rPr")

    rFonts = rPr.find(f"{{{W}}}rFonts")
    if rFonts is None:
        rFonts = etree.SubElement(rPr, f"{{{W}}}rFonts")
    rFonts.set(f"{{{W}}}ascii", "Times New Roman")
    rFonts.set(f"{{{W}}}hAnsi", "Times New Roman")
    rFonts.set(f"{{{W}}}eastAsia", settings.title_font)

    half_points = str(int(size_label_to_points(settings.title_size) * 2))
    for tag in ("sz", "szCs"):
        elem = rPr.find(f"{{{W}}}{tag}")
        if elem is None:
            elem = etree.SubElement(rPr, f"{{{W}}}{tag}")
        elem.set(_W_VAL, half_points)

    for tag in ("b", "bCs"):
        _set_bool_property(rPr, tag, settings.title_bold)


def _format_paragraph_mark_for_caption(para, settings) -> None:
    """Format the paragraph mark so caption fields inherit the desired style."""
    pPr = para._element.find(f"{{{W}}}pPr")
    if pPr is None:
        pPr = etree.SubElement(para._element, f"{{{W}}}pPr")
    rPr = pPr.find(f"{{{W}}}rPr")
    if rPr is None:
        rPr = etree.SubElement(pPr, f"{{{W}}}rPr")
    for tag in ("b", "bCs"):
        _set_bool_property(rPr, tag, settings.title_bold)


def apply_table_caption(para, settings, doc=None) -> None:
    """Apply formatting to a table caption paragraph."""
    # Normalize space between number and caption text: "表 1-1  测试" → "表 1-1 测试"
    _normalize_caption_space(para)
    format_all_runs_in_paragraph(
        para,
        ascii_font="Times New Roman",
        eastasia_font=settings.title_font,
        size_pt=size_label_to_points(settings.title_size),
        bold=settings.title_bold,
    )
    _format_paragraph_mark_for_caption(para, settings)
    _format_numbering_level_for_caption(para, doc, settings)

    # Remove any outline level — captions are not headings
    _remove_outline_level(para)

    _ALIGN_MAP = {"左对齐": "left", "居中": "center", "右对齐": "right", "两端对齐": "both"}
    set_paragraph_alignment(para, _ALIGN_MAP.get(settings.title_alignment, "center"))

    # Apply caption indentation from settings
    apply_indent_cm(
        para.paragraph_format,
        left_cm=settings.title_left_indent_cm,
        right_cm=settings.title_right_indent_cm,
        special_kind=settings.title_special_indent,
        special_cm=settings.title_special_indent_cm,
    )

    apply_paragraph_spacing(
        para.paragraph_format,
        before_lines=settings.title_before_lines,
        after_lines=settings.title_after_lines,
        line_spacing=settings.title_spacing,
    )


def apply_table_body(table, settings) -> None:
    """Apply formatting to table body."""
    # Table centering
    set_table_alignment(table, "center")

    # Column layout mode
    set_table_layout(table, settings.auto_fit_mode)

    # Borders
    set_table_borders(table, line_width_pt=settings.line_width_pt)

    if settings.row_height_cm > 0:
        for row in table.rows:
            set_row_height_at_least(row, settings.row_height_cm)

    # Cell font — first row (header) bold, rest normal
    for row_idx, row in enumerate(table.rows):
        bold_row = (row_idx == 0)
        row_spacing = 1.0 if bold_row else settings.spacing  # 单倍 for header row
        for cell in row.cells:
            for para in cell.paragraphs:
                format_all_runs_in_paragraph(
                    para,
                    ascii_font="Times New Roman",
                    eastasia_font=settings.body_font,
                    size_pt=size_label_to_points(settings.body_size),
                    bold=bold_row,
                )
                apply_paragraph_spacing(
                    para.paragraph_format,
                    line_spacing=row_spacing,
                )
                apply_indent_cm(
                    para.paragraph_format,
                    left_cm=0.0,
                    right_cm=0.0,
                    special_kind="无",
                    special_chars=0.0,
                    special_cm=0.0,
                )


def _normalize_caption_space(para) -> None:
    """Ensure exactly one space between caption number and title text.

    "表 1-1    测试" or "表 1.8.12-2\t测试" → one separator space.

    Only normalizes when the full caption is within a single run,
    to avoid corrupting multi-run paragraphs.
    """
    import re
    for run in para.runs:
        m = re.match(rf"^([表图]\s*{_CAPTION_NUMBER_PATTERN})(\s{{2,}})(.+)$", run.text)
        if m:
            run.text = m.group(1) + ' ' + m.group(3)
            return


def refresh_all(doc, settings, *, _paragraphs=None, _tables=None) -> list[str]:
    """Refresh all tables and their captions.

    Returns a list of warning messages for captions found in shapes/text frames
    that could not be formatted.
    """
    tables = _tables if _tables is not None else doc.tables
    formatted_captions = set()
    for table in tables:
        caption = find_table_caption(table, doc, _tables=tables, _paragraphs=_paragraphs)
        if caption is not None:
            apply_table_caption(caption, settings, doc)
            formatted_captions.add(id(caption._element))
        apply_table_body(table, settings)

    # Also format standalone table captions (paragraphs matching the caption
    # pattern that were not found by find_table_caption — e.g. captions not
    # immediately preceding a table element).
    paragraphs = _paragraphs if _paragraphs is not None else doc.paragraphs
    for para in paragraphs:
        if is_table_caption_paragraph(para, doc) and id(para._element) not in formatted_captions:
            apply_table_caption(para, settings, doc)

    # One-time scan for captions in shapes/text frames.
    # Scan body elements in order, collecting shape caption texts with their
    # body position. Then match each shape caption to the nearest following table.
    warnings: list[str] = _collect_shape_caption_warnings(doc, tables)
    return warnings


def _collect_shape_caption_warnings(doc, tables: list) -> list[str]:
    """Scan shapes/text frames once and map captions to the correct table by position."""
    from lxml import etree
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    A = "http://schemas.openxmlformats.org/drawingml/2006/main"

    # Collect (caption_text, body_element_index) for shape captions
    shape_captions: list[tuple[str, int]] = []
    body_elements = list(doc.element.body)

    for i, elem in enumerate(body_elements):
        if not elem.tag.endswith("}p"):
            continue
        for drawing in elem.findall(f".//{{{W}}}drawing"):
            tx_body = drawing.find(f".//{{{A}}}txBody")
            if tx_body is None:
                continue
            for ap in tx_body.findall(f"{{{A}}}p"):
                para_text = ""
                for r in ap.findall(f".//{{{A}}}t"):
                    if r.text:
                        para_text += r.text
                if is_table_caption_line(para_text):
                    shape_captions.append((para_text[:80], i))
                    break  # One caption per drawing is enough
            else:
                continue
            break

    if not shape_captions:
        return []

    # Map each shape caption to the nearest following table
    # Collect (body_element_index, table_index) for all tables
    table_positions: list[tuple[int, int]] = []
    for ti, table in enumerate(tables):
        for j, elem in enumerate(body_elements):
            if elem is table._element:
                table_positions.append((j, ti))
                break

    warnings = []
    for caption_text, shape_pos in shape_captions:
        # Find nearest following table
        nearest_table_idx = None
        for table_pos, ti in table_positions:
            if table_pos > shape_pos:
                nearest_table_idx = ti
                break
        if nearest_table_idx is not None:
            warnings.append(
                f"表格{nearest_table_idx + 1}的表题在文本框/形状中（\"{caption_text}...\"），"
                "当前版本无法修改形状内文本，请手动格式化。"
            )
    return warnings

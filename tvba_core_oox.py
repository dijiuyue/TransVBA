"""OOXML helpers using lxml direct element manipulation.

Thin wrappers that look like python-docx API but operate at the lxml level
for attributes that python-docx does not expose.
"""
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _ns(tag: str) -> str:
    return f"{{{W}}}{tag}"


def _ensure_rPr(run) -> etree._Element:
    rPr = run._element.find(_ns("rPr"))
    if rPr is None:
        rPr = etree.SubElement(run._element, _ns("rPr"))
    return rPr


def _ensure_pPr(para) -> etree._Element:
    pPr = para._element.find(_ns("pPr"))
    if pPr is None:
        pPr = etree.SubElement(para._element, _ns("pPr"))
    return pPr


def set_far_east_font(run, font_name: str) -> None:
    """Set East Asian font via w:rFonts/@w:eastAsia."""
    rPr = _ensure_rPr(run)
    rFonts = rPr.find(_ns("rFonts"))
    if rFonts is None:
        rFonts = etree.SubElement(rPr, _ns("rFonts"))
    rFonts.set(_ns("eastAsia"), font_name)


def set_ascii_font(run, font_name: str) -> None:
    """Set ASCII font via python-docx run.font.name."""
    run.font.name = font_name


def set_run_font_size(run, points: float) -> None:
    """Set font size on a run in both w:sz (western) and w:szCs (eastAsian).

    python-docx run.font.size only sets w:sz. For CJK documents, Word also
    reads w:szCs for east Asian text, so both must be set.
    """
    rPr = _ensure_rPr(run)
    half_points = str(int(points * 2))
    sz = rPr.find(_ns("sz"))
    if sz is None:
        sz = etree.SubElement(rPr, _ns("sz"))
    sz.set(_ns("val"), half_points)
    szCs = rPr.find(_ns("szCs"))
    if szCs is None:
        szCs = etree.SubElement(rPr, _ns("szCs"))
    szCs.set(_ns("val"), half_points)


def set_style_font_size(style, points: float) -> None:
    """Set font size on a style in both w:sz and w:szCs.

    python-docx style.font.size only sets w:sz.
    """
    rPr = style.element.find(_ns("rPr"))
    if rPr is None:
        rPr = etree.SubElement(style.element, _ns("rPr"))
    half_points = str(int(points * 2))
    sz = rPr.find(_ns("sz"))
    if sz is None:
        sz = etree.SubElement(rPr, _ns("sz"))
    sz.set(_ns("val"), half_points)
    szCs = rPr.find(_ns("szCs"))
    if szCs is None:
        szCs = etree.SubElement(rPr, _ns("szCs"))
    szCs.set(_ns("val"), half_points)


def set_outline_level(paragraph, level_zero_indexed: int) -> None:
    """Set paragraph outline level (0-8)."""
    pPr = _ensure_pPr(paragraph)
    outline = pPr.find(_ns("outlineLvl"))
    if outline is None:
        outline = etree.SubElement(pPr, _ns("outlineLvl"))
    outline.set(_ns("val"), str(level_zero_indexed))


def apply_indent_chars(
    paragraph_format,
    *,
    left_chars: float,
    right_chars: float,
    special_kind: str,
    special_chars: float,
) -> None:
    """Apply indentation in character units (1 char = 12 pt = 240 twips).

    special_kind: "无", "首行缩进", "悬挂缩进"
    """
    # paragraph_format._element is the w:p element; we need w:pPr
    p_elem = paragraph_format._element
    pPr = p_elem.find(_ns("pPr"))
    if pPr is None:
        pPr = etree.SubElement(p_elem, _ns("pPr"))
    ind = pPr.find(_ns("ind"))
    if ind is None:
        ind = etree.SubElement(pPr, _ns("ind"))

    twips_per_char = 240  # 12 pt * 20 twips/pt

    if left_chars:
        ind.set(_ns("left"), str(int(left_chars * twips_per_char)))
    if right_chars:
        ind.set(_ns("right"), str(int(right_chars * twips_per_char)))

    # Clear existing special indent attrs
    for attr in (_ns("firstLine"), _ns("hanging")):
        if attr in ind.attrib:
            del ind.attrib[attr]

    if special_kind == "首行缩进" and special_chars:
        ind.set(_ns("firstLine"), str(int(special_chars * twips_per_char)))
    elif special_kind == "悬挂缩进" and special_chars:
        ind.set(_ns("hanging"), str(int(special_chars * twips_per_char)))


def set_before_after_lines(paragraph_format, *, before_lines: float, after_lines: float) -> None:
    """Set paragraph spacing in line units (w:beforeLines / w:afterLines).

    1 line = 100 hundredths of a line (same as VBA LineUnitBefore).
    """
    # paragraph_format._element is the w:p element; we need w:pPr
    p_elem = paragraph_format._element
    pPr = p_elem.find(_ns("pPr"))
    if pPr is None:
        pPr = etree.SubElement(p_elem, _ns("pPr"))
    spacing = pPr.find(_ns("spacing"))
    if spacing is None:
        spacing = etree.SubElement(pPr, _ns("spacing"))
    spacing.set(_ns("beforeLines"), str(int(before_lines * 100)))
    spacing.set(_ns("afterLines"), str(int(after_lines * 100)))


from tvba_utils import cm_to_points


def set_table_layout_window(table) -> None:
    """Set table to autofit to window (AutoFitBehavior=2)."""
    tblPr = table._element.find(_ns("tblPr"))
    if tblPr is None:
        tblPr = etree.SubElement(table._element, _ns("tblPr"))
    layout = tblPr.find(_ns("tblLayout"))
    if layout is None:
        layout = etree.SubElement(tblPr, _ns("tblLayout"))
    layout.set(_ns("val"), "autofit")


def set_table_layout_content(table) -> None:
    """Set table to autofit to content (AutoFitBehavior=1)."""
    tblPr = table._element.find(_ns("tblPr"))
    if tblPr is None:
        tblPr = etree.SubElement(table._element, _ns("tblPr"))
    layout = tblPr.find(_ns("tblLayout"))
    if layout is None:
        layout = etree.SubElement(tblPr, _ns("tblLayout"))
    layout.set(_ns("val"), "fixed")


def set_table_borders(table, *, line_width_pt: float) -> None:
    """Set all table borders to a uniform width."""
    tblPr = table._element.find(_ns("tblPr"))
    if tblPr is None:
        tblPr = etree.SubElement(table._element, _ns("tblPr"))
    borders = tblPr.find(_ns("tblBorders"))
    if borders is None:
        borders = etree.SubElement(tblPr, _ns("tblBorders"))
    sz = str(int(line_width_pt * 20))  # half-points (1 pt = 20 half-points)
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = borders.find(_ns(side))
        if border is None:
            border = etree.SubElement(borders, _ns(side))
        border.set(_ns("val"), "single")
        border.set(_ns("sz"), sz)
        border.set(_ns("space"), "0")
        border.set(_ns("color"), "auto")


def set_row_height_at_least(row, height_cm: float) -> None:
    """Set row height to at least the given cm."""
    trPr = row._tr.find(_ns("trPr"))
    if trPr is None:
        trPr = etree.SubElement(row._tr, _ns("trPr"))
    trHeight = trPr.find(_ns("trHeight"))
    if trHeight is None:
        trHeight = etree.SubElement(trPr, _ns("trHeight"))
    points = cm_to_points(height_cm)
    twips = int(points * 20)
    trHeight.set(_ns("val"), str(twips))
    trHeight.set(_ns("hRule"), "atLeast")

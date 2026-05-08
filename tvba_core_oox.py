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

"""Body text formatting.

Corresponds to VBA FormatModule.bas:
  - RefreshContentFormat (BodyText branch)
  - ApplySettingsToDocument (wdStyleNormal setting)
"""
from tvba_core_oox import (
    set_far_east_font,
    set_ascii_font,
    apply_indent_chars,
    set_before_after_lines,
)
from tvba_utils import size_label_to_points, cm_to_points
from docx.shared import Pt

_ALIGNMENT_MAP = {
    "左对齐": 0,
    "居中": 1,
    "右对齐": 2,
    "两端对齐": 3,
}

def apply_normal_style(doc, body) -> None:
    """Apply body settings to the Normal style."""
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(size_label_to_points(body.size))
    # East Asian font
    for para in doc.paragraphs:
        for run in para.runs:
            set_far_east_font(run, body.font)
        break  # Only need to set on style, but python-docx style font lacks eastAsia
    # Set line spacing on Normal style via direct XML on default pPr
    pPr = normal.element.find(".//w:pPr", {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
    if pPr is not None:
        from lxml import etree
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        spacing = pPr.find(f"{{{W}}}spacing")
        if spacing is None:
            spacing = etree.SubElement(pPr, f"{{{W}}}spacing")
        spacing.set(f"{{{W}}}line", str(int(body.spacing * 240)))
        spacing.set(f"{{{W}}}lineRule", "auto")

def apply_paragraph(para, body) -> None:
    """Apply body formatting to a single paragraph."""
    # Font on each run
    for run in para.runs:
        set_ascii_font(run, "Times New Roman")
        set_far_east_font(run, body.font)
        run.font.size = Pt(size_label_to_points(body.size))

    # Alignment
    para.alignment = _ALIGNMENT_MAP.get(body.alignment, 3)

    # Indent
    apply_indent_chars(
        para.paragraph_format,
        left_chars=0.0,
        right_chars=0.0,
        special_kind=body.special_indent,
        special_chars=cm_to_points(body.special_indent_cm) / 12.0,  # convert pt to chars
    )

    # Spacing (before/after in lines)
    set_before_after_lines(
        para.paragraph_format,
        before_lines=body.before_lines,
        after_lines=body.after_lines,
    )

    # Line spacing
    pPr = para._element.find(".//w:pPr", {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
    if pPr is not None:
        from lxml import etree
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        spacing = pPr.find(f"{{{W}}}spacing")
        if spacing is None:
            spacing = etree.SubElement(pPr, f"{{{W}}}spacing")
        spacing.set(f"{{{W}}}line", str(int(body.spacing * 240)))
        spacing.set(f"{{{W}}}lineRule", "auto")

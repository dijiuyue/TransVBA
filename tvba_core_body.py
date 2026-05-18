"""Body text formatting.

Corresponds to VBA FormatModule.bas:
  - RefreshContentFormat (BodyText branch)
  - ApplySettingsToDocument (wdStyleNormal setting)
"""
from tvba_core_oox import (
    set_far_east_font,
    set_ascii_font,
    set_run_font_size,
    set_style_font_size,
    apply_indent_chars,
    apply_indent_cm,
    apply_paragraph_spacing,
)
from tvba_core_normalize import apply_brackets, add_period_if_needed
from tvba_utils import size_label_to_points, cm_to_points

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
    set_style_font_size(normal, size_label_to_points(body.size))
    # East Asian font
    for para in doc.paragraphs:
        for run in para.runs:
            set_far_east_font(run, body.font)
        break  # Only need to set on style, but python-docx style font lacks eastAsia
    # Line spacing is set per-paragraph in apply_paragraph, not on Normal style,
    # to avoid stretching empty paragraphs that serve as visual separators.

def apply_paragraph(para, body) -> None:
    """Apply body formatting to a single paragraph."""
    # Font on each run
    for run in para.runs:
        set_ascii_font(run, "Times New Roman")
        set_far_east_font(run, body.font)
        set_run_font_size(run, size_label_to_points(body.size))

    # Alignment
    para.alignment = _ALIGNMENT_MAP.get(body.alignment, 3)

    # Indent — left/right in cm, special in chars
    apply_indent_cm(
        para.paragraph_format,
        left_cm=body.left_indent_cm,
        right_cm=body.right_indent_cm,
        special_kind=body.special_indent,
        special_chars=body.special_indent_chars,
    )

    # Spacing (before/after in lines, line spacing)
    apply_paragraph_spacing(
        para.paragraph_format,
        before_lines=body.before_lines,
        after_lines=body.after_lines,
        line_spacing=body.spacing,
    )

    # Content modification (brackets, periods) only when explicitly enabled —
    # formatting should not silently change text.
    if body.modify_content:
        text = para.text
        if text:
            apply_brackets(para, text)
            add_period_if_needed(para)

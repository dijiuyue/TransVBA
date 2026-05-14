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
    set_before_after_lines,
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

    # Indent
    apply_indent_chars(
        para.paragraph_format,
        left_chars=0.0,
        right_chars=0.0,
        special_kind=body.special_indent,
        special_chars=body.special_indent_chars,
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

    # Normalize brackets, add period, fix forbidden words
    text = para.text
    if text:
        apply_brackets(para, text)
        add_period_if_needed(para)
        # Skip forbidden-word replacement for figure/table captions
        if not _is_caption_line(text):
            _replace_forbidden_words(para)


_FORBIDDEN_MAP = {"附图": "附件", "附表": "附件"}
_CAPTION_LINE_RE = __import__('re').compile(r'^(?:[表图]\s*\d+(?:\.\d+)*-\d+\s|附图?\s*\d+(?:\.\d+)*\s)')


def _is_caption_line(text: str) -> bool:
    """Check if text looks like a figure/table caption (should skip forbidden-word replacement)."""
    return bool(_CAPTION_LINE_RE.match(text.strip()))


def _replace_forbidden_words(para) -> None:
    """Replace forbidden words in paragraph runs."""
    for run in para.runs:
        text = run.text
        if not text:
            continue
        for bad, good in _FORBIDDEN_MAP.items():
            if bad in text:
                run.text = text.replace(bad, good)

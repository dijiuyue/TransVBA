"""Figure caption formatting.

Corresponds to VBA FormatModule.bas:
  - RefreshFigureCaptions
  - IsFigureCaptionLine
"""
from tvba_core_oox import set_before_after_lines, apply_indent_chars, format_all_runs_in_paragraph
import re

from tvba_utils import clean_para_text, size_label_to_points

# VBA pattern: ^图\s*\d+(\.\d+)*-\d+[\t ]+.+$
# Use \s+ to match any whitespace (space, tab, nbsp, fullwidth space, etc.)
_FIGURE_CAPTION_RE = re.compile(
    r"^(?:图|figure)\s*\d+(?:\.\d+)*-\d+\s+.+$",
    re.IGNORECASE,
)


def is_figure_caption_line(text: str) -> bool:
    """Check if text is a figure caption."""
    text = clean_para_text(text)
    return bool(_FIGURE_CAPTION_RE.match(text))


def apply_figure_caption(para, settings) -> None:
    """Apply formatting to a figure caption paragraph."""
    format_all_runs_in_paragraph(
        para,
        ascii_font="Times New Roman",
        eastasia_font=settings.title_font,
        size_pt=size_label_to_points(settings.title_size),
        bold=settings.title_bold,
    )

    para.alignment = 1  # Center

    # Clear any inherited indentation (body first-line/hanging indent)
    apply_indent_chars(
        para.paragraph_format,
        left_chars=0.0,
        right_chars=0.0,
        special_kind="无",
        special_chars=0.0,
    )

    set_before_after_lines(
        para.paragraph_format,
        before_lines=0.0,
        after_lines=0.0,
    )

    # Line spacing
    pPr = para._element.find(".//w:pPr", {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
    if pPr is not None:
        from lxml import etree
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        spacing = pPr.find(f"{{{W}}}spacing")
        if spacing is None:
            spacing = etree.SubElement(pPr, f"{{{W}}}spacing")
        spacing.set(f"{{{W}}}line", str(int(settings.title_spacing * 240)))
        spacing.set(f"{{{W}}}lineRule", "auto")


def refresh_all(doc, settings) -> None:
    """Refresh all figure captions in document."""
    for para in doc.paragraphs:
        if is_figure_caption_line(para.text):
            apply_figure_caption(para, settings)

"""TOC detection and styling.

Corresponds to VBA FormatModule.bas:
  - RefreshDirectoryFormat
  - IsTocEntryLine
  - IsTocParagraph
  - IsDirectoryTitleLine
  - IdentifyDirectoryLevel
  - ApplyTocStyleToParagraph
  - ApplyDirectoryTitleStyle
  - ApplyDirectoryStyle
"""
from tvba_core_oox import set_far_east_font, set_ascii_font
from tvba_utils import clean_para_text, size_label_to_points
from docx.shared import Pt


def is_toc_entry_line(text: str) -> bool:
    """Check if text is a TOC entry: contains Tab and last token is numeric."""
    text = clean_para_text(text)
    if "\t" not in text:
        return False
    parts = text.split("\t")
    # Last non-empty part should be a number
    last = parts[-1].strip()
    if not last:
        return False
    # Remove trailing \r if any
    last = last.replace("\r", "").strip()
    try:
        float(last)
        return True
    except ValueError:
        return False


def is_toc_title_line(text: str) -> bool:
    """Check if text is the TOC title ('目录')."""
    return clean_para_text(text) == "目录"


def is_toc_paragraph(para) -> bool:
    """Check if paragraph is part of TOC (entry or title)."""
    text = clean_para_text(para.text)
    return is_toc_entry_line(text) or is_toc_title_line(text)


def identify_toc_level(text: str) -> int:
    """Identify TOC entry level from leading whitespace."""
    # Do NOT strip leading spaces — they indicate indent level
    text = text.rstrip().replace("\r", "").replace("\n", "")
    if not text.startswith(" "):
        return 1
    # Count leading spaces
    stripped = text.lstrip(" ")
    spaces = len(text) - len(stripped)
    if spaces == 2:
        return 2
    if spaces == 4:
        return 3
    return 0


def apply_toc_title_style(para, defaults) -> None:
    """Apply TOC title formatting."""
    for run in para.runs:
        set_ascii_font(run, "Times New Roman")
        set_far_east_font(run, defaults.title_font)
        run.font.size = Pt(size_label_to_points(defaults.title_size))
        run.font.bold = defaults.title_bold
    # Line spacing
    pPr = para._element.find(".//w:pPr", {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
    if pPr is not None:
        from lxml import etree
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        spacing = pPr.find(f"{{{W}}}spacing")
        if spacing is None:
            spacing = etree.SubElement(pPr, f"{{{W}}}spacing")
        spacing.set(f"{{{W}}}line", str(int(defaults.title_spacing * 240)))
        spacing.set(f"{{{W}}}lineRule", "auto")


def apply_toc_entry_style(doc, para, level: int, defaults) -> None:
    """Apply TOC entry formatting with style + direct format override."""
    style_name = f"TOC {level}"
    try:
        para.style = doc.styles[style_name]
    except KeyError:
        pass  # Style may not exist

    if level == 1:
        font = defaults.level1_font
        size = defaults.level1_size
        bold = defaults.level1_bold
    elif level == 2:
        font = defaults.level2_font
        size = defaults.level2_size
    elif level == 3:
        font = defaults.level3_font
        size = defaults.level3_size
    else:
        font = defaults.level1_font
        size = defaults.level1_size
        bold = False

    for run in para.runs:
        set_ascii_font(run, "Times New Roman")
        set_far_east_font(run, font)
        run.font.size = Pt(size_label_to_points(size))
        if level == 1:
            run.font.bold = bold


def refresh_toc(doc, defaults) -> None:
    """Refresh all TOC paragraphs in document."""
    for para in doc.paragraphs:
        text = clean_para_text(para.text)
        if is_toc_title_line(text):
            apply_toc_title_style(para, defaults)
        elif is_toc_entry_line(text):
            level = identify_toc_level(text)
            if level == 0:
                level = 1
            apply_toc_entry_style(doc, para, level, defaults)

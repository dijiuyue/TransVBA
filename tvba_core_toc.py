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
import re

from tvba_core_oox import apply_paragraph_spacing, set_far_east_font, set_ascii_font, set_run_font_size
from tvba_utils import clean_para_text, size_label_to_points

# Matches numeric prefix like "1", "1.1", "1.1.2" (halfwidth or fullwidth digits/dots)
_TOC_NUMBER_RE = re.compile(r"^([\d０-９]+([\.．][\d０-９]+){0,6})[ \t]*")

# Matches TOC title: "目录" or "目  录" (characters separated by any whitespace)
_TOC_TITLE_RE = re.compile(r"^目\s+录$")

# Matches space-separated TOC entries: number + 2+ spaces + text + 2+ spaces + page number
_TOC_ENTRY_SPACE_RE = re.compile(
    r"^[\d０-９]+[\.．\d０-９]*\s{2,}.+\s{2,}\d+$"
)


def is_toc_entry_line(text: str) -> bool:
    """Check if text is a TOC entry: contains Tab and last token is numeric."""
    text = clean_para_text(text)
    # Tab-separated entries (standard Word TOC format)
    if "\t" in text:
        parts = text.split("\t")
        last = parts[-1].strip()
        if not last:
            return False
        last = last.replace("\r", "").strip()
        try:
            float(last)
            return True
        except ValueError:
            return False
    # Space-separated entries: number + 2+ spaces + text + 2+ spaces + page number
    if _TOC_ENTRY_SPACE_RE.match(text):
        return True
    return False


def is_toc_title_line(text: str) -> bool:
    """Check if text is the TOC title ('目录' or '目  录' with spacing)."""
    text = clean_para_text(text)
    return text == "目录" or bool(_TOC_TITLE_RE.match(text))


def _get_para_style_id(para) -> str | None:
    """Fast path: read paragraph style id directly from XML."""
    from lxml import etree
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    pPr = para._element.find(f".//{{{W}}}pPr")
    if pPr is None:
        return None
    pStyle = pPr.find(f"{{{W}}}pStyle")
    if pStyle is None:
        return None
    return pStyle.get(f"{{{W}}}val")


def is_toc_paragraph(para, _toc_style_ids=None) -> bool:
    """Check if paragraph is part of TOC (entry or title).

    Returns True if the text looks like a TOC entry/title, OR if the
    paragraph's style id contains "TOC" (case-insensitive).

    Optional _toc_style_ids is a set of style ids whose style names contain
    "toc" — provided as a fast-path alternative to expensive paragraph.style lookup.
    """
    text = clean_para_text(para.text)
    if is_toc_entry_line(text) or is_toc_title_line(text):
        return True
    # Fast path: read style id directly from XML to avoid expensive
    # python-docx paragraph.style lookup which traverses the entire styles collection.
    style_id = _get_para_style_id(para)
    if style_id and "toc" in style_id.lower():
        return True
    # Fallback: check against pre-built set of TOC style ids (covers cases
    # where style id doesn't contain "toc" but style name does, e.g. custom styles).
    if _toc_style_ids and style_id in _toc_style_ids:
        return True
    return False


def identify_toc_level(text: str) -> int:
    """Identify TOC entry level from number prefix or leading whitespace.

    Priority:
      1. Extract number prefix (e.g. "1", "1.1", "1.1.2") and count dots.
      2. Fall back to leading space counting if no number pattern matches.
    """
    # Priority 1: regex-based number prefix detection
    m = _TOC_NUMBER_RE.match(text)
    if m:
        num_part = m.group(1)
        # Normalize fullwidth digits/dots to halfwidth
        for i in range(10):
            num_part = num_part.replace(chr(0xFF10 + i), str(i))
        num_part = num_part.replace("．", ".")
        dot_count = num_part.count(".")
        level = dot_count + 1
        if 1 <= level <= 3:
            return level
        return 0

    # Priority 2: fallback to leading space counting
    text = text.rstrip().replace("\r", "").replace("\n", "")
    if not text.startswith(" "):
        return 1
    stripped = text.lstrip(" ")
    spaces = len(text) - len(stripped)
    if spaces == 2:
        return 2
    if spaces == 4:
        return 3
    return 0


def apply_toc_title_style(para, defaults) -> None:
    """Apply TOC title formatting."""
    para.alignment = 1  # Center
    for run in para.runs:
        set_ascii_font(run, "Times New Roman")
        set_far_east_font(run, defaults.title_font)
        set_run_font_size(run, size_label_to_points(defaults.title_size))
        run.font.bold = defaults.title_bold
    apply_paragraph_spacing(
        para.paragraph_format,
        line_spacing=defaults.title_spacing,
    )


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
        set_run_font_size(run, size_label_to_points(size))
        if level == 1:
            run.font.bold = bold


def refresh_toc(doc, defaults, *, _paragraphs=None) -> None:
    """Refresh all TOC paragraphs in document.

    Also ensures exactly one blank line after the TOC title.
    """
    from lxml import etree
    W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    paragraphs = _paragraphs if _paragraphs is not None else doc.paragraphs
    for i, para in enumerate(paragraphs):
        text = clean_para_text(para.text)
        if is_toc_title_line(text):
            apply_toc_title_style(para, defaults)
            # Ensure one blank line after TOC title
            _ensure_one_blank_after(doc, paragraphs, i, para._element, W_NS)
        elif is_toc_entry_line(text):
            level = identify_toc_level(text)
            if level == 0:
                level = 1
            apply_toc_entry_style(doc, para, level, defaults)


def _ensure_one_blank_after(doc, paragraphs: list, idx: int, para_elem, W_NS: str) -> None:
    """Ensure exactly one blank paragraph after the given paragraph.

    Removes extra empty paragraphs and inserts one if none exists.
    """
    from lxml import etree
    body = doc.element.body
    body_children = list(body)
    try:
        elem_idx = body_children.index(para_elem)
    except ValueError:
        return

    # Find consecutive empty paragraphs after this element
    to_remove = []
    for j in range(elem_idx + 1, len(body_children)):
        next_elem = body_children[j]
        tag = next_elem.tag.split("}")[-1] if "}" in next_elem.tag else next_elem.tag
        if tag != "p":
            break
        # Check if empty
        texts = []
        for t in next_elem.findall(f".//{{{W_NS}}}t"):
            if t.text:
                texts.append(t.text)
        if not "".join(texts).strip():
            to_remove.append(next_elem)
        else:
            break

    if len(to_remove) > 1:
        # Remove extras, keep one
        for extra in to_remove[1:]:
            body.remove(extra)
    elif len(to_remove) == 0:
        # Insert one blank paragraph
        new_p = etree.SubElement(body, f"{{{W_NS}}}p")
        # Move it right after the TOC title
        body.remove(new_p)
        body.insert(elem_idx + 1, new_p)

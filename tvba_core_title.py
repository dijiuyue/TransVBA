"""Title detection and formatting.

Corresponds to VBA FormatModule.bas:
  - AutoDetectAndFormatNumericTitles (line 783-826)
  - IdentifyContentTitleLevel (line 840-893)
  - IdentifyContentTitleLevelFromNumber (line 901-949)
  - ApplyContentTitleStyle (line 951+)
  - NormalizeNumberString (line 829-837)
"""
import re

from tvba_core_oox import (
    set_far_east_font,
    set_ascii_font,
    set_run_font_size,
    set_outline_level,
    get_effective_outline_level,
    apply_indent_chars,
    set_before_after_lines,
)
from tvba_core_normalize import (
    apply_brackets,
    add_period_if_needed,
    sync_number_font_with_body,
)
from tvba_utils import size_label_to_points, clean_para_text

# Matches numeric title prefix: digits (and fullwidth digits) with dots/fullwidth dots,
# optionally ending with a trailing dot, followed by space/tab and title text.
_TITLE_RE = re.compile(r"^([\d０-９]+([\.．]\d+){0,6}\.?)[ \t]+(.+)$")


def normalize_number_string(s: str) -> str:
    """Normalize number string: fullwidth digits/dots to halfwidth, strip trailing dot, trim."""
    s = s.strip()
    # Fullwidth digits ０-９ (U+FF10-U+FF19) to halfwidth 0-9
    for i in range(10):
        s = s.replace(chr(0xFF10 + i), str(i))
    s = s.replace("．", ".")
    s = s.replace("。", ".")
    if s.endswith("."):
        s = s[:-1]
    return s


def identify_level_from_number(num_str: str) -> int:
    """Map a normalized number string to title level 1-5 (0 = not a title)."""
    if not num_str:
        return 0
    dot_count = num_str.count(".")
    if dot_count == 0:
        return 1
    if dot_count == 1 and num_str.endswith(".0"):
        return 1
    level = dot_count + 1
    if 1 <= level <= 5:
        return level
    return 0


def identify_numeric_title_level(text: str) -> int:
    """Identify title level from paragraph text. Returns 1-5 or 0."""
    text = clean_para_text(text)
    m = _TITLE_RE.match(text)
    if not m:
        return 0
    num_part = normalize_number_string(m.group(1))
    return identify_level_from_number(num_part)


def apply_title_style(paragraph, level: int, level_settings, body_settings) -> None:
    """Apply title formatting to a paragraph."""
    # Set outline level (0-indexed: level 1 -> 0)
    set_outline_level(paragraph, level - 1)

    # Font on each run
    for run in paragraph.runs:
        set_ascii_font(run, "Times New Roman")
        set_far_east_font(run, level_settings.font)
        set_run_font_size(run, size_label_to_points(level_settings.size))
        run.font.bold = level_settings.bold

    # Alignment
    _ALIGNMENT_MAP = {"左对齐": 0, "居中": 1, "右对齐": 2, "两端对齐": 3}
    paragraph.alignment = _ALIGNMENT_MAP.get(level_settings.alignment, 0)

    # Spacing
    set_before_after_lines(
        paragraph.paragraph_format,
        before_lines=level_settings.before_lines,
        after_lines=level_settings.after_lines,
    )

    # Line spacing
    pPr = paragraph._element.find(".//w:pPr", {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
    if pPr is not None:
        from lxml import etree
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        spacing = pPr.find(f"{{{W}}}spacing")
        if spacing is None:
            spacing = etree.SubElement(pPr, f"{{{W}}}spacing")
        spacing.set(f"{{{W}}}line", str(int(level_settings.line_spacing * 240)))
        spacing.set(f"{{{W}}}lineRule", "auto")

    # Normalize brackets, add period, sync number font
    apply_brackets(paragraph, paragraph.text)
    add_period_if_needed(paragraph)
    sync_number_font_with_body(paragraph)


def auto_detect_and_format(doc, settings, list_resolver=None) -> None:
    """Auto-detect numeric titles and apply title formatting."""
    from tvba_core_toc import is_toc_paragraph

    for para in doc.paragraphs:
        if is_toc_paragraph(para):
            continue

        # Skip paragraphs that already have an outline level set by the user
        # (either directly on the paragraph or inherited from its style)
        existing_outline = get_effective_outline_level(para, styles=doc.styles)
        if existing_outline is not None:
            continue

        text = clean_para_text(para.text)
        if not text:
            continue

        # Priority 1: Numeric title text detection
        level = identify_numeric_title_level(text)

        # Priority 2: Multi-level list resolver (COM or docx) as fallback.
        # Only trust list resolver when it can prove the paragraph is a
        # heading (via list text for COM) to avoid treating body list
        # items (e.g. "1）第一项") as titles.
        if level == 0 and list_resolver is not None:
            list_level = list_resolver.get_list_level(para)
            if list_level is not None and 1 <= list_level <= 5:
                # COM resolver: verify via rendered list text
                if hasattr(list_resolver, 'get_list_text'):
                    list_text = list_resolver.get_list_text(para)
                    if list_text:
                        normalized = normalize_number_string(list_text)
                        if identify_level_from_number(normalized) > 0:
                            level = list_level
                # Docx resolver: cannot reliably distinguish heading lists
                # from body lists, so skip to avoid false positives.
                # Text already matched above if it looks like a title.

        if 1 <= level <= 5:
            apply_title_style(
                para,
                level,
                settings.titles[level - 1],
                settings.body,
            )

"""Cover page title detection and formatting.

Detects the cover page title on the first page and applies formatting
from the active template's cover settings.
Only targets short centered paragraphs that look like a document title
(not headings, TOC entries, or body text).
"""
import re
from tvba_core_oox import format_all_runs_in_paragraph, get_effective_outline_level
from tvba_core_toc import is_toc_title_line, is_toc_entry_line
from tvba_settings import CoverSettings
from tvba_utils import size_label_to_points

_ALIGNMENT_MAP = {"居中": 1, "左对齐": 0, "右对齐": 2, "两端对齐": 3}


def format_cover_title(doc, settings: "CoverSettings | None" = None, *, _paragraphs=None) -> None:
    """Detect and format the cover page title.

    Only targets the first qualifying centered paragraph that:
    - Has no outline level (not a heading)
    - Is not part of the TOC
    - Is reasonably short (like a title, not body text)
    - Appears before any TOC entries (i.e., on the cover page)
    """
    if settings is None:
        settings = CoverSettings()
    paragraphs = _paragraphs if _paragraphs is not None else list(doc.paragraphs)

    alignment_val = _ALIGNMENT_MAP.get(settings.alignment, 1)

    for para in paragraphs[:settings.search_paragraphs]:
        text = (para.text or "").strip()
        if not text or len(text) < 2 or len(text) > settings.max_length:
            continue
        if re.match(r'^[\d\.\s]+$', text):
            continue

        # Stop at TOC boundary — cover title is always before the TOC
        if is_toc_title_line(text) or is_toc_entry_line(text):
            break

        if para.alignment != alignment_val:
            continue

        # Skip paragraphs with outline levels (headings)
        outline = get_effective_outline_level(para)
        if outline is not None:
            continue

        # Found a likely cover title — format it and stop
        format_all_runs_in_paragraph(
            para,
            ascii_font="Times New Roman",
            eastasia_font=settings.font,
            size_pt=size_label_to_points(settings.size),
            bold=settings.bold,
        )
        para.alignment = alignment_val

        from lxml import etree
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        pPr = para._element.find(f".//{{{W}}}pPr")
        if pPr is None:
            pPr = etree.SubElement(para._element, f"{{{W}}}pPr")
        spacing = pPr.find(f"{{{W}}}spacing")
        if spacing is None:
            spacing = etree.SubElement(pPr, f"{{{W}}}spacing")
        spacing.set(f"{{{W}}}line", str(int(settings.line_spacing * 240)))
        spacing.set(f"{{{W}}}lineRule", "auto")
        break

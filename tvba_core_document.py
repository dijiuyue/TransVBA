"""Document orchestrator.

Corresponds to VBA FormatModule.bas:
  - ApplySettingsToDocument
"""
from pathlib import Path

from docx import Document

from tvba_settings import FormatSettings
from tvba_core_body import apply_normal_style, apply_paragraph
from tvba_core_title import auto_detect_and_format, split_compound_paragraphs
from tvba_core_toc import is_toc_paragraph, refresh_toc
from tvba_core_table import refresh_all as refresh_tables
from tvba_core_figure import refresh_all as refresh_figures
from tvba_core_convert import ensure_docx
from tvba_core_normalize import unify_ascii_font
from tvba_core_numbering import auto_select
from tvba_core_oox import get_effective_outline_level, sync_numbering_with_titles
from tvba_core_cover import format_cover_title
from tvba_core_appendix import format_appendix


def apply_settings_to_document(
    docx_path: Path,
    settings: FormatSettings,
    *,
    list_resolver=None,
    output_path: Path | None = None,
    progress_cb=None,
) -> Path:
    """Apply all formatting settings to a document.

    Returns the output path (output_path or docx_path for in-place).
    """
    if progress_cb:
        progress_cb("Loading document...", 0.0)

    docx_path = ensure_docx(docx_path)

    doc = Document(str(docx_path))

    # Cache paragraph/table lists to avoid repeated rebuilding.
    # python-docx rebuilds these from XML on every access.
    paragraphs = list(doc.paragraphs)
    tables = list(doc.tables)

    # Cache outline level lookups — style chain resolution is expensive
    # and the same paragraph may be checked multiple times.
    outline_cache: dict[int, int | None] = {}

    # Pre-build style-by-id map because python-docx deprecated style_id lookup.
    style_by_id = {style.style_id: style for style in doc.styles}

    # Pre-build TOC style id set: style ids whose names contain "toc".
    # This lets is_toc_paragraph detect custom TOC styles without the expensive
    # paragraph.style lookup.
    toc_style_ids = {style.style_id for style in doc.styles if "toc" in style.name.lower()}

    if progress_cb:
        progress_cb("Applying normal style...", 0.1)
    apply_normal_style(doc, settings.body)

    # Pre-process: split compound paragraphs with multiple concatenated titles
    if progress_cb:
        progress_cb("Splitting compound titles...", 0.15)
    paragraphs = split_compound_paragraphs(doc, _paragraphs=paragraphs)

    # Auto-detect titles
    if progress_cb:
        progress_cb("Detecting titles...", 0.2)
    if settings.auto_detect_numeric_titles:
        if list_resolver is None and settings.auto_detect_include_list_paragraphs:
            list_resolver = auto_select(prefer_com=settings.prefer_com_resolver, docx_path=str(docx_path), doc=doc)
        auto_detect_and_format(doc, settings, list_resolver, _paragraphs=paragraphs, _outline_cache=outline_cache, _style_by_id=style_by_id, _toc_style_ids=toc_style_ids)

    if progress_cb:
        progress_cb("Formatting paragraphs...", 0.4)
    for para in paragraphs:
        if is_toc_paragraph(para, _toc_style_ids=toc_style_ids):
            continue

        # Skip empty paragraphs — keep their original spacing/layout
        if not para.text or not para.text.strip():
            continue

        # Check effective outline level (direct or from style)
        outline_level = get_effective_outline_level(para, _cache=outline_cache, _style_by_id=style_by_id)

        if outline_level is not None and 0 <= outline_level <= 4:
            # Title paragraph (Word levels 1-5) — apply title formatting
            from tvba_core_title import apply_title_style
            level = outline_level + 1  # Convert 0-4 to 1-5
            apply_title_style(para, level, settings.titles[level - 1], settings.body)
        else:
            # Body text (no outline or level > 4) — apply body formatting
            apply_paragraph(para, settings.body)

    # Sync numbering definitions so auto-generated list numbers match titles
    sync_numbering_with_titles(doc, settings, _paragraphs=paragraphs)

    if progress_cb:
        progress_cb("Formatting TOC...", 0.6)
    refresh_toc(doc, settings.toc, _paragraphs=paragraphs)

    if progress_cb:
        progress_cb("Formatting tables...", 0.7)
    refresh_tables(doc, settings.table, _paragraphs=paragraphs, _tables=tables)

    if progress_cb:
        progress_cb("Formatting figures...", 0.8)
    refresh_figures(doc, settings.figure, _paragraphs=paragraphs)

    if progress_cb:
        progress_cb("Formatting cover page...", 0.82)
    format_cover_title(doc, settings.cover, _paragraphs=paragraphs)

    if progress_cb:
        progress_cb("Formatting appendix...", 0.84)
    format_appendix(doc, settings.appendix, _paragraphs=paragraphs, _outline_cache=outline_cache, _style_by_id=style_by_id)

    if progress_cb:
        progress_cb("Formatting headers...", 0.86)
    _format_headers(doc, settings.header)

    if progress_cb:
        progress_cb("Normalizing fonts...", 0.9)
    unify_ascii_font(doc, "Times New Roman", _paragraphs=paragraphs, _tables=tables)

    if progress_cb:
        progress_cb("Saving...", 0.95)

    # Close COM resolver before saving to release file lock
    if list_resolver is not None and hasattr(list_resolver, "close"):
        list_resolver.close()

    out = output_path or docx_path
    doc.save(str(out))

    if progress_cb:
        progress_cb("Done", 1.0)

    return out


def _format_headers(doc, settings) -> None:
    """Format header text in all sections using settings values.

    Also normalize 'Rev.' spacing: ensure single space after 'Rev.'
    """
    from tvba_core_oox import format_all_runs_in_paragraph
    from tvba_utils import size_label_to_points

    for section in doc.sections:
        header = section.header
        if header is None:
            continue
        for para in header.paragraphs:
            format_all_runs_in_paragraph(
                para,
                ascii_font="Times New Roman",
                eastasia_font=settings.font,
                size_pt=size_label_to_points(settings.size),
                bold=settings.bold,
            )
            # Line spacing
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

            # Normalize "Rev." spacing: "Rev.  " → "Rev. "
            import re
            for run in header.paragraphs[0].runs if header.paragraphs else []:
                if 'Rev.' in run.text:
                    run.text = re.sub(r'Rev\.\s+', 'Rev. ', run.text)

"""Document orchestrator.

Corresponds to VBA FormatModule.bas:
  - ApplySettingsToDocument
"""
from pathlib import Path

from docx import Document

from tvba_settings import FormatSettings
from tvba_core_body import apply_normal_style, apply_paragraph
from tvba_core_title import auto_detect_and_format
from tvba_core_toc import is_toc_paragraph, refresh_toc
from tvba_core_table import refresh_all as refresh_tables
from tvba_core_figure import refresh_all as refresh_figures
from tvba_core_convert import ensure_docx
from tvba_core_normalize import unify_ascii_font
from tvba_core_numbering import auto_select


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

    if progress_cb:
        progress_cb("Applying normal style...", 0.1)
    apply_normal_style(doc, settings.body)

    # Auto-detect titles
    if progress_cb:
        progress_cb("Detecting titles...", 0.2)
    if settings.auto_detect_numeric_titles:
        if list_resolver is None and settings.auto_detect_include_list_paragraphs:
            list_resolver = auto_select(prefer_com=settings.prefer_com_resolver, docx_path=str(docx_path), doc=doc)
        auto_detect_and_format(doc, settings, list_resolver)

    if progress_cb:
        progress_cb("Formatting paragraphs...", 0.4)
    for para in doc.paragraphs:
        if is_toc_paragraph(para):
            continue

        # Check if paragraph has outline level set (title)
        pPr = para._element.find(".//w:pPr", {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
        outline_level = None
        if pPr is not None:
            outline = pPr.find("w:outlineLvl", {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
            if outline is not None:
                outline_level = int(outline.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "9"))

        if outline_level is not None and 0 <= outline_level <= 4:
            # Title paragraph (Word levels 1-5) — apply title formatting
            from tvba_core_title import apply_title_style
            level = outline_level + 1  # Convert 0-4 to 1-5
            apply_title_style(para, level, settings.titles[level - 1], settings.body)
        else:
            # Body text (outline_level == 9 or no outline) — apply body formatting
            apply_paragraph(para, settings.body)

    if progress_cb:
        progress_cb("Formatting TOC...", 0.6)
    refresh_toc(doc, settings.toc)

    if progress_cb:
        progress_cb("Formatting tables...", 0.7)
    refresh_tables(doc, settings.table)

    if progress_cb:
        progress_cb("Formatting figures...", 0.8)
    refresh_figures(doc, settings.figure)

    if progress_cb:
        progress_cb("Normalizing fonts...", 0.9)
    unify_ascii_font(doc, "Times New Roman")

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

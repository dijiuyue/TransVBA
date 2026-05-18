"""Appendix title and body formatting.

Detects appendix titles ("附件...") and applies formatting from
the active template's appendix settings.
"""
import re
from tvba_core_oox import (
    apply_paragraph_spacing,
    format_all_runs_in_paragraph,
    get_effective_outline_level,
)
from tvba_settings import AppendixSettings
from tvba_utils import size_label_to_points


_APPENDIX_TITLE_RE = re.compile(r'^附件\s*\d*')


def is_appendix_title(text: str) -> bool:
    """Check if text is an appendix title line."""
    return bool(_APPENDIX_TITLE_RE.match(text.strip()))


def format_appendix(doc, settings: "AppendixSettings | None" = None, *, _paragraphs=None, _outline_cache=None, _style_by_id=None) -> None:
    """Format appendix titles and body text in the document."""
    if settings is None:
        settings = AppendixSettings()
    paragraphs = _paragraphs if _paragraphs is not None else list(doc.paragraphs)

    in_appendix = False
    for para in paragraphs:
        text = (para.text or "").strip()
        if not text:
            continue

        # Check if this paragraph is a heading (ends appendix body zone)
        outline = get_effective_outline_level(para, _cache=_outline_cache, _style_by_id=_style_by_id)
        if outline is not None and outline <= 4:
            in_appendix = False

        if is_appendix_title(text):
            in_appendix = True
            _format_appendix_title(para, text, settings)
            continue

        if in_appendix:
            _format_appendix_body(para, settings)


def _format_appendix_title(para, text: str, settings: AppendixSettings) -> None:
    """Format an appendix title using settings values."""
    for run in para.runs:
        m = _APPENDIX_TITLE_RE.match(run.text)
        if m:
            prefix_end = m.end()
            rest = run.text[prefix_end:]
            if rest and rest[0] not in ('：', ':'):
                run.text = run.text[:prefix_end] + '：' + rest
            elif not rest:
                run.text = run.text[:prefix_end] + '：'
            break

    format_all_runs_in_paragraph(
        para,
        ascii_font="Times New Roman",
        eastasia_font=settings.title_font,
        size_pt=size_label_to_points(settings.title_size),
        bold=settings.title_bold,
    )
    para.alignment = 0  # left

    apply_paragraph_spacing(
        para.paragraph_format,
        line_spacing=settings.title_line_spacing,
    )


def _format_appendix_body(para, settings: AppendixSettings) -> None:
    """Format appendix body text using settings values."""
    format_all_runs_in_paragraph(
        para,
        ascii_font="Times New Roman",
        eastasia_font=settings.body_font,
        size_pt=size_label_to_points(settings.body_size),
        bold=settings.body_bold,
    )

    apply_paragraph_spacing(
        para.paragraph_format,
        line_spacing=settings.body_line_spacing,
    )

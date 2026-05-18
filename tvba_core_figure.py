"""Figure caption formatting.

Corresponds to VBA FormatModule.bas:
  - RefreshFigureCaptions
  - IsFigureCaptionLine
"""
from tvba_core_oox import apply_paragraph_spacing, apply_indent_cm, set_paragraph_alignment, format_all_runs_in_paragraph
import re
from lxml import etree

from tvba_utils import clean_para_text, size_label_to_points

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W_VAL = f"{{{W}}}val"

# Caption number category:
#   图1 标题
#   图1-1 标题
#   图1.8.3-1 标题
# Also accepts full-width digits/dots/dashes copied from Word/WPS.
_CAPTION_NUMBER_PATTERN = r"[0-9０-９]+(?:[.．][0-9０-９]+)*(?:[-‑－–—\x1e][0-9０-９]+)?"
_FIGURE_CAPTION_RE = re.compile(
    rf"^(?:图|figure)\s*{_CAPTION_NUMBER_PATTERN}\s*[^\s0-9０-９.．\-‑－–—\x1e].*$",
    re.IGNORECASE,
)


def is_figure_caption_line(text: str) -> bool:
    """Check if text is a figure caption."""
    text = clean_para_text(text)
    return bool(_FIGURE_CAPTION_RE.match(text))


def is_figure_caption_paragraph(para, doc=None) -> bool:
    """Check whether a paragraph is a figure caption.

    Real Word documents may store the visible "图2.5.2-1" prefix as an
    automatic numbering label rather than as paragraph text.  In that case the
    paragraph text is only the caption title ("作业带布置示意图"), while
    w:numPr points to a numbering level whose w:lvlText starts with "图".
    """
    text = clean_para_text(para.text)
    if is_figure_caption_line(text):
        return True
    if not text:
        return False
    if _style_is_figure_caption(para, doc):
        return True
    return _numbering_label_starts_with_figure(para, doc)


def _style_is_figure_caption(para, doc=None) -> bool:
    """Return True for paragraph styles whose Word style name starts with 图."""
    style_id = _paragraph_style_id(para)
    if not style_id:
        return False
    style = _style_element_by_id(doc, style_id)
    if style is None:
        return False
    name = style.find(f"{{{W}}}name")
    style_name = name.get(_W_VAL, "") if name is not None else ""
    return bool(re.match(rf"^图\s*{_CAPTION_NUMBER_PATTERN}", style_name))


def _numbering_label_starts_with_figure(para, doc=None) -> bool:
    """Return True if paragraph numbering format starts with 图."""
    ilvl, num_id = _effective_num_pr(para, doc)
    if ilvl is None or num_id is None:
        return False
    lvl_text = _numbering_level_text(doc, num_id, ilvl)
    return bool(lvl_text and lvl_text.lstrip().startswith("图"))


def _paragraph_style_id(para) -> str | None:
    pPr = para._element.find(f"{{{W}}}pPr")
    if pPr is None:
        return None
    pStyle = pPr.find(f"{{{W}}}pStyle")
    return None if pStyle is None else pStyle.get(_W_VAL)


def _effective_num_pr(para, doc=None) -> tuple[int | None, str | None]:
    """Read direct numbering first, then inherit numbering from paragraph style."""
    pPr = para._element.find(f"{{{W}}}pPr")
    if pPr is not None:
        num_pr = pPr.find(f"{{{W}}}numPr")
        ilvl, num_id = _num_pr_values(num_pr)
        if num_id is not None:
            return ilvl, num_id

    style_id = _paragraph_style_id(para)
    seen: set[str] = set()
    while style_id and style_id not in seen:
        seen.add(style_id)
        style = _style_element_by_id(doc, style_id)
        if style is None:
            break
        num_pr = style.find(f"{{{W}}}pPr/{{{W}}}numPr")
        ilvl, num_id = _num_pr_values(num_pr)
        if num_id is not None:
            if ilvl is None:
                outline = style.find(f"{{{W}}}pPr/{{{W}}}outlineLvl")
                ilvl = _int_or_none(outline.get(_W_VAL)) if outline is not None else 0
            return ilvl, num_id
        based_on = style.find(f"{{{W}}}basedOn")
        style_id = based_on.get(_W_VAL) if based_on is not None else None
    return None, None


def _num_pr_values(num_pr) -> tuple[int | None, str | None]:
    if num_pr is None:
        return None, None
    ilvl_elem = num_pr.find(f"{{{W}}}ilvl")
    num_id_elem = num_pr.find(f"{{{W}}}numId")
    ilvl = _int_or_none(ilvl_elem.get(_W_VAL)) if ilvl_elem is not None else None
    num_id = num_id_elem.get(_W_VAL) if num_id_elem is not None else None
    return ilvl, num_id


def _style_element_by_id(doc, style_id: str):
    styles_element = _styles_element(doc)
    if styles_element is None:
        return None
    return styles_element.find(f".//{{{W}}}style[@{{{W}}}styleId='{style_id}']")


def _styles_element(doc):
    if doc is None:
        return None
    try:
        return doc.styles.element
    except Exception:
        return None


def _numbering_level_text(doc, num_id: str, ilvl: int | None) -> str | None:
    lvl = _numbering_level(doc, num_id, ilvl)
    return _lvl_text(lvl)


def _numbering_level(doc, num_id: str, ilvl: int | None):
    numbering = _numbering_element(doc)
    if numbering is None or ilvl is None:
        return None

    num = numbering.find(f".//{{{W}}}num[@{{{W}}}numId='{num_id}']")
    if num is None:
        return None

    # A level override may carry a full level definition. Prefer it when present.
    override = num.find(f"{{{W}}}lvlOverride[@{{{W}}}ilvl='{ilvl}']")
    if override is not None:
        lvl = override.find(f"{{{W}}}lvl")
        if lvl is not None:
            return lvl

    abstract_id = num.find(f"{{{W}}}abstractNumId")
    if abstract_id is None:
        return None
    abstract = numbering.find(f".//{{{W}}}abstractNum[@{{{W}}}abstractNumId='{abstract_id.get(_W_VAL)}']")
    if abstract is None:
        return None
    return abstract.find(f"{{{W}}}lvl[@{{{W}}}ilvl='{ilvl}']")


def _lvl_text(lvl) -> str | None:
    if lvl is None:
        return None
    lvl_text = lvl.find(f"{{{W}}}lvlText")
    return None if lvl_text is None else lvl_text.get(_W_VAL)


def _numbering_element(doc):
    if doc is None:
        return None
    try:
        return doc.part.numbering_part._element
    except Exception:
        return None


def _int_or_none(value) -> int | None:
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _remove_outline_level(para) -> None:
    """Remove any outlineLvl element so captions are not treated as headings."""
    pPr = para._element.find(f"{{{W}}}pPr")
    if pPr is not None:
        outline = pPr.find(f"{{{W}}}outlineLvl")
        if outline is not None:
            pPr.remove(outline)


def _set_bool_property(rPr, tag: str, value: bool) -> None:
    elem = rPr.find(f"{{{W}}}{tag}")
    if elem is None:
        elem = etree.SubElement(rPr, f"{{{W}}}{tag}")
    if value:
        elem.attrib.pop(_W_VAL, None)
    else:
        elem.set(_W_VAL, "0")


def _set_paragraph_mark_bold(para, bold: bool) -> None:
    pPr = para._element.find(f"{{{W}}}pPr")
    if pPr is None:
        pPr = etree.SubElement(para._element, f"{{{W}}}pPr")
    rPr = pPr.find(f"{{{W}}}rPr")
    if rPr is None:
        rPr = etree.SubElement(pPr, f"{{{W}}}rPr")
    for tag in ("b", "bCs"):
        _set_bool_property(rPr, tag, bold)


def _set_numbering_level_bold(para, doc, bold: bool) -> None:
    ilvl, num_id = _effective_num_pr(para, doc)
    lvl = _numbering_level(doc, num_id, ilvl) if num_id is not None else None
    if lvl is None:
        return
    rPr = lvl.find(f"{{{W}}}rPr")
    if rPr is None:
        rPr = etree.SubElement(lvl, f"{{{W}}}rPr")
    for tag in ("b", "bCs"):
        _set_bool_property(rPr, tag, bold)


def apply_figure_caption(para, settings, doc=None) -> None:
    """Apply formatting to a figure caption paragraph."""
    _normalize_caption_space(para)
    format_all_runs_in_paragraph(
        para,
        ascii_font="Times New Roman",
        eastasia_font=settings.title_font,
        size_pt=size_label_to_points(settings.title_size),
        bold=settings.title_bold,
    )
    _set_paragraph_mark_bold(para, settings.title_bold)
    _set_numbering_level_bold(para, doc, settings.title_bold)

    # Remove any outline level — captions are not headings
    _remove_outline_level(para)

    _ALIGN_MAP = {"左对齐": "left", "居中": "center", "右对齐": "right", "两端对齐": "both"}
    set_paragraph_alignment(para, _ALIGN_MAP.get(settings.title_alignment, "center"))

    # Apply caption indentation from settings
    apply_indent_cm(
        para.paragraph_format,
        left_cm=settings.title_left_indent_cm,
        right_cm=settings.title_right_indent_cm,
        special_kind=settings.title_special_indent,
        special_cm=settings.title_special_indent_cm,
    )

    apply_paragraph_spacing(
        para.paragraph_format,
        before_lines=settings.title_before_lines,
        after_lines=settings.title_after_lines,
        line_spacing=settings.title_spacing,
    )


def _normalize_caption_space(para) -> None:
    """Ensure exactly one space between caption number and title text."""
    from tvba_core_table import _normalize_caption_space as _impl
    _impl(para)


def refresh_all(doc, settings, *, _paragraphs=None) -> None:
    """Refresh all figure captions in document."""
    paragraphs = _paragraphs if _paragraphs is not None else doc.paragraphs
    for para in paragraphs:
        if is_figure_caption_paragraph(para, doc):
            apply_figure_caption(para, settings, doc)

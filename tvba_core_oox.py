"""OOXML helpers using lxml direct element manipulation.

Thin wrappers that look like python-docx API but operate at the lxml level
for attributes that python-docx does not expose.
"""
from dataclasses import dataclass
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


@dataclass(frozen=True)
class ParagraphFormatSpec:
    """Unified paragraph formatting specification.

    Replaces scattered per-module formatting logic with a single spec
    that can be applied consistently across body, titles, captions, and headers.
    """
    eastasia_font: str
    ascii_font: str = "Times New Roman"
    size_pt: float = 12.0
    bold: bool = False
    alignment: str | None = None  # "左对齐", "居中", "右对齐", "两端对齐", or None to skip
    before_lines: float = 0.0
    after_lines: float = 0.0
    line_spacing: float | None = None  # None to skip
    left_chars: float = 0.0
    right_chars: float = 0.0
    special_kind: str = "无"  # "无", "首行缩进", "悬挂缩进"
    special_chars: float = 0.0
    outline_level: int | None = None  # 0-8 (0-indexed), None to skip

    def apply_to(self, para) -> None:
        """Apply this spec to a paragraph, formatting all runs."""
        format_all_runs_in_paragraph(
            para,
            ascii_font=self.ascii_font,
            eastasia_font=self.eastasia_font,
            size_pt=self.size_pt,
            bold=self.bold,
        )

        if self.alignment is not None:
            _ALIGN_MAP = {"左对齐": 0, "居中": 1, "右对齐": 2, "两端对齐": 3}
            para.alignment = _ALIGN_MAP.get(self.alignment, 0)

        apply_indent_chars(
            para.paragraph_format,
            left_chars=self.left_chars,
            right_chars=self.right_chars,
            special_kind=self.special_kind,
            special_chars=self.special_chars,
        )

        apply_paragraph_spacing(
            para.paragraph_format,
            before_lines=self.before_lines,
            after_lines=self.after_lines,
            line_spacing=self.line_spacing,
        )

        if self.outline_level is not None:
            set_outline_level(para, self.outline_level)


def _ns(tag: str) -> str:
    return f"{{{W}}}{tag}"


def _ensure_rPr(run) -> etree._Element:
    rPr = run._element.find(_ns("rPr"))
    if rPr is None:
        rPr = etree.SubElement(run._element, _ns("rPr"))
    return rPr


def _ensure_pPr(para) -> etree._Element:
    pPr = para._element.find(_ns("pPr"))
    if pPr is None:
        pPr = etree.SubElement(para._element, _ns("pPr"))
    return pPr


def format_all_runs_in_paragraph(para, *, ascii_font: str, eastasia_font: str, size_pt: float, bold: bool = False) -> None:
    """Format all runs in a paragraph, including nested ones (fields, hyperlinks).

    python-docx's para.runs only returns direct children. Runs inside
    w:hyperlink, w:fldSimple, etc. are missed. This function walks the
    entire paragraph XML tree and formats every w:r element.
    """
    half_points = str(int(size_pt * 2))
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    for run_elem in para._element.findall(f".//{{{W}}}r"):
        rPr = run_elem.find(f"{{{W}}}rPr")
        if rPr is None:
            rPr = etree.SubElement(run_elem, f"{{{W}}}rPr")

        # Set fonts
        rFonts = rPr.find(f"{{{W}}}rFonts")
        if rFonts is None:
            rFonts = etree.SubElement(rPr, f"{{{W}}}rFonts")
        rFonts.set(f"{{{W}}}ascii", ascii_font)
        rFonts.set(f"{{{W}}}hAnsi", ascii_font)
        rFonts.set(f"{{{W}}}eastAsia", eastasia_font)

        # Set size (both western and CJK)
        sz = rPr.find(f"{{{W}}}sz")
        if sz is None:
            sz = etree.SubElement(rPr, f"{{{W}}}sz")
        sz.set(f"{{{W}}}val", half_points)

        szCs = rPr.find(f"{{{W}}}szCs")
        if szCs is None:
            szCs = etree.SubElement(rPr, f"{{{W}}}szCs")
        szCs.set(f"{{{W}}}val", half_points)

        # Set bold explicitly. Removing w:b is not enough when the paragraph
        # style or numbering level is bold; w:val="0" overrides inheritance.
        # Normalize duplicate properties from Word-authored documents.
        for tag in ("b", "bCs"):
            for old in list(rPr.findall(f"{{{W}}}{tag}")):
                rPr.remove(old)
            b = etree.SubElement(rPr, f"{{{W}}}{tag}")
            if bold:
                b.attrib.pop(f"{{{W}}}val", None)
            else:
                b.set(f"{{{W}}}val", "0")


def clear_paragraph_formatting(para) -> None:
    """Remove run-level formatting (font, size, bold) from all runs in a paragraph.

    Preserves text content — only clears formatting to prepare for fresh styling.
    Also clears paragraph-level spacing and indentation.
    """
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    for run_elem in para._element.findall(f".//{{{W}}}r"):
        rPr = run_elem.find(f"{{{W}}}rPr")
        if rPr is None:
            continue

        for tag in ("rFonts", "sz", "szCs", "b", "bCs", "i", "iCs", "u", "color", "highlight", "spacing", "kern"):
            for elem in list(rPr.findall(f"{{{W}}}{tag}")):
                rPr.remove(elem)

    # Clear paragraph spacing
    pPr = para._element.find(f"{{{W}}}pPr")
    if pPr is not None:
        for tag in ("spacing", "ind", "jc"):
            elem = pPr.find(f"{{{W}}}{tag}")
            if elem is not None:
                pPr.remove(elem)


def set_far_east_font(run, font_name: str) -> None:
    """Set East Asian font via w:rFonts/@w:eastAsia."""
    rPr = _ensure_rPr(run)
    rFonts = rPr.find(_ns("rFonts"))
    if rFonts is None:
        rFonts = etree.SubElement(rPr, _ns("rFonts"))
    rFonts.set(_ns("eastAsia"), font_name)


def set_ascii_font(run, font_name: str) -> None:
    """Set ASCII font via python-docx run.font.name."""
    run.font.name = font_name


def set_run_font_size(run, points: float) -> None:
    """Set font size on a run in both w:sz (western) and w:szCs (eastAsian).

    python-docx run.font.size only sets w:sz. For CJK documents, Word also
    reads w:szCs for east Asian text, so both must be set.
    """
    rPr = _ensure_rPr(run)
    half_points = str(int(points * 2))
    sz = rPr.find(_ns("sz"))
    if sz is None:
        sz = etree.SubElement(rPr, _ns("sz"))
    sz.set(_ns("val"), half_points)
    szCs = rPr.find(_ns("szCs"))
    if szCs is None:
        szCs = etree.SubElement(rPr, _ns("szCs"))
    szCs.set(_ns("val"), half_points)


def set_style_font_size(style, points: float) -> None:
    """Set font size on a style in both w:sz and w:szCs.

    python-docx style.font.size only sets w:sz.
    """
    rPr = style.element.find(_ns("rPr"))
    if rPr is None:
        rPr = etree.SubElement(style.element, _ns("rPr"))
    half_points = str(int(points * 2))
    sz = rPr.find(_ns("sz"))
    if sz is None:
        sz = etree.SubElement(rPr, _ns("sz"))
    sz.set(_ns("val"), half_points)
    szCs = rPr.find(_ns("szCs"))
    if szCs is None:
        szCs = etree.SubElement(rPr, _ns("szCs"))
    szCs.set(_ns("val"), half_points)


def set_outline_level(paragraph, level_zero_indexed: int) -> None:
    """Set paragraph outline level (0-8)."""
    pPr = _ensure_pPr(paragraph)
    outline = pPr.find(_ns("outlineLvl"))
    if outline is None:
        outline = etree.SubElement(pPr, _ns("outlineLvl"))
    outline.set(_ns("val"), str(level_zero_indexed))


def _get_para_style_id(para) -> str | None:
    """Read paragraph style id directly from XML (fast path avoiding python-docx style lookup)."""
    pPr = para._element.find(_ns("pPr"))
    if pPr is None:
        return None
    pStyle = pPr.find(_ns("pStyle"))
    if pStyle is None:
        return None
    return pStyle.get(_ns("val"))


def get_effective_outline_level(paragraph, styles=None, _cache=None, _style_by_id=None) -> int | None:
    """Return the effective outline level (0-8) for a paragraph.

    Checks the paragraph's direct w:pPr/w:outlineLvl first, then falls back
    to the paragraph style (and its basedOn chain). Returns None if no
    outline level is found.

    Optional _cache dict can be passed to memoize results by paragraph element id.
    _style_by_id is a pre-built dict mapping style_id -> Style to avoid deprecated
    python-docx style_id lookup.
    """
    para_id = id(paragraph._element)
    if _cache is not None and para_id in _cache:
        return _cache[para_id]

    # 1. Direct paragraph-level outline
    pPr = paragraph._element.find(_ns("pPr"))
    if pPr is not None:
        outline = pPr.find(_ns("outlineLvl"))
        if outline is not None:
            val = outline.get(_ns("val"))
            if val is not None:
                result = int(val)
                if _cache is not None:
                    _cache[para_id] = result
                return result

    # 2. Style-level outline (follow basedOn chain)
    # Fast path: read style id from XML directly instead of using paragraph.style
    # which triggers an expensive O(n) lookup in the styles collection.
    style_id = _get_para_style_id(paragraph)
    style = None
    if style_id is not None:
        if _style_by_id is not None:
            style = _style_by_id.get(style_id)
        elif styles is not None:
            # Deprecated fallback — may trigger warnings in newer python-docx
            try:
                style = styles[style_id]
            except KeyError:
                style = None
        else:
            style = paragraph.style

    _styles_lookup = _style_by_id if _style_by_id is not None else styles

    while style is not None:
        style_pPr = style.element.find(_ns("pPr"))
        if style_pPr is not None:
            outline = style_pPr.find(_ns("outlineLvl"))
            if outline is not None:
                val = outline.get(_ns("val"))
                if val is not None:
                    result = int(val)
                    if _cache is not None:
                        _cache[para_id] = result
                    return result

        # Follow basedOn chain
        basedOn = style.element.find(_ns("basedOn"))
        if basedOn is None:
            break
        based_on_val = basedOn.get(_ns("val"))
        if based_on_val is None or _styles_lookup is None:
            break
        try:
            style = _styles_lookup[based_on_val]
        except KeyError:
            break

    if _cache is not None:
        _cache[para_id] = None
    return None


def apply_indent_chars(
    paragraph_format,
    *,
    left_chars: float,
    right_chars: float,
    special_kind: str,
    special_chars: float,
) -> None:
    """Apply indentation in character units (1 char = 12 pt = 240 twips).

    special_kind: "无", "首行缩进", "悬挂缩进"
    """
    p_elem = paragraph_format._element
    pPr = p_elem.find(_ns("pPr"))
    if pPr is None:
        pPr = etree.SubElement(p_elem, _ns("pPr"))
    ind = pPr.find(_ns("ind"))
    if ind is None:
        ind = etree.SubElement(pPr, _ns("ind"))

    if left_chars:
        ind.set(_ns("leftChars"), str(int(left_chars * 100)))
        ind.attrib.pop(_ns("left"), None)
    else:
        ind.set(_ns("leftChars"), "0")
        ind.attrib.pop(_ns("left"), None)

    if right_chars:
        ind.set(_ns("rightChars"), str(int(right_chars * 100)))
        ind.attrib.pop(_ns("right"), None)
    else:
        ind.set(_ns("rightChars"), "0")
        ind.attrib.pop(_ns("right"), None)

    # Clear existing special indent attrs
    for attr in (_ns("firstLine"), _ns("hanging"), _ns("firstLineChars"), _ns("hangingChars")):
        if attr in ind.attrib:
            del ind.attrib[attr]

    if special_kind == "首行缩进" and special_chars:
        ind.set(_ns("firstLineChars"), str(int(special_chars * 100)))
    elif special_kind == "悬挂缩进" and special_chars:
        ind.set(_ns("hangingChars"), str(int(special_chars * 100)))
    else:
        ind.set(_ns("firstLineChars"), "0")
        ind.set(_ns("firstLine"), "0")


def apply_style_indent_chars(
    style,
    *,
    left_chars: float,
    right_chars: float,
    special_kind: str,
    special_chars: float,
) -> None:
    pPr = style.element.find(_ns("pPr"))
    if pPr is None:
        pPr = etree.SubElement(style.element, _ns("pPr"))
    ind = pPr.find(_ns("ind"))
    if ind is None:
        ind = etree.SubElement(pPr, _ns("ind"))

    for attr in (_ns("left"), _ns("right"), _ns("firstLine"), _ns("hanging"), _ns("firstLineChars"), _ns("hangingChars")):
        if attr in ind.attrib:
            del ind.attrib[attr]

    ind.set(_ns("leftChars"), str(int(left_chars * 100)))
    ind.set(_ns("rightChars"), str(int(right_chars * 100)))
    if special_kind == "\u9996\u884c\u7f29\u8fdb" and special_chars:
        ind.set(_ns("firstLineChars"), str(int(special_chars * 100)))
    elif special_kind == "\u60ac\u6302\u7f29\u8fdb" and special_chars:
        ind.set(_ns("hangingChars"), str(int(special_chars * 100)))
    else:
        ind.set(_ns("firstLineChars"), "0")
        ind.set(_ns("firstLine"), "0")


def apply_indent_cm(
    paragraph_format,
    *,
    left_cm: float = 0.0,
    right_cm: float = 0.0,
    special_kind: str = "无",
    special_chars: float = 0.0,
    special_cm: float = 0.0,
) -> None:
    """Apply indentation with cm-based left/right/special and char-based special.

    Left/right indentation in cm. Special indent in either cm or chars.
    If both special_cm and special_chars are given, special_cm takes precedence.
    1 cm = 567 twips, 1 char = 240 twips (12 pt).
    """
    p_elem = paragraph_format._element
    pPr = p_elem.find(_ns("pPr"))
    if pPr is None:
        pPr = etree.SubElement(p_elem, _ns("pPr"))
    ind = pPr.find(_ns("ind"))
    if ind is None:
        ind = etree.SubElement(pPr, _ns("ind"))

    twips_per_cm = 567
    twips_per_char = 240

    if left_cm:
        ind.set(_ns("left"), str(int(left_cm * twips_per_cm)))
    else:
        ind.set(_ns("left"), "0")

    if right_cm:
        ind.set(_ns("right"), str(int(right_cm * twips_per_cm)))
    else:
        ind.set(_ns("right"), "0")

    for attr in (_ns("firstLine"), _ns("hanging"), _ns("firstLineChars"), _ns("hangingChars")):
        if attr in ind.attrib:
            del ind.attrib[attr]

    if special_kind == "首行缩进":
        if special_cm:
            ind.set(_ns("firstLine"), str(int(special_cm * twips_per_cm)))
        elif special_chars:
            ind.set(_ns("firstLine"), str(int(special_chars * twips_per_char)))
    elif special_kind == "悬挂缩进":
        if special_cm:
            ind.set(_ns("hanging"), str(int(special_cm * twips_per_cm)))
        elif special_chars:
            ind.set(_ns("hanging"), str(int(special_chars * twips_per_char)))


def set_before_after_lines(paragraph_format, *, before_lines: float, after_lines: float) -> None:
    """Apply paragraph spacing via the unified helper (legacy wrapper)."""
    apply_paragraph_spacing(paragraph_format, before_lines=before_lines, after_lines=after_lines)


def apply_paragraph_spacing(
    paragraph_format,
    *,
    before_lines: float = 0.0,
    after_lines: float = 0.0,
    line_spacing: float | None = None,
    line_rule: str = "auto",
) -> None:
    """Unified paragraph spacing helper.

    - Removes ALL old spacing attributes that could interfere:
      w:before, w:after, w:beforeLines, w:afterLines,
      w:beforeAutospacing, w:afterAutospacing
    - Writes explicit beforeLines / afterLines in line units.
    - When before/after is 0, also writes w:before="0" / w:after="0"
      to prevent inherited style spacing from overriding.
    - Optionally sets line spacing (w:line / w:lineRule).
    """
    p_elem = paragraph_format._element
    pPr = p_elem.find(_ns("pPr"))
    if pPr is None:
        pPr = etree.SubElement(p_elem, _ns("pPr"))

    spacing = pPr.find(_ns("spacing"))

    # Remove ALL old spacing attributes — any of these can interfere
    # with the template-defined spacing values.
    dirty_attrs = (
        "before", "after", "beforeLines", "afterLines",
        "beforeAutospacing", "afterAutospacing",
    )
    if spacing is not None:
        for attr in dirty_attrs:
            key = _ns(attr)
            if key in spacing.attrib:
                del spacing.attrib[key]

    # Write explicit before/after in point units (w:before/w:after) to
    # prevent inherited style spacing from overriding when targeting 0.
    if spacing is None:
        spacing = etree.SubElement(pPr, _ns("spacing"))
    spacing.set(_ns("before"), "0" if before_lines == 0 else str(int(before_lines * 240)))
    spacing.set(_ns("after"), "0" if after_lines == 0 else str(int(after_lines * 240)))

    # Write beforeLines / afterLines as primary spacing control
    spacing.set(_ns("beforeLines"), str(int(before_lines * 100)))
    spacing.set(_ns("afterLines"), str(int(after_lines * 100)))

    # Disable auto-spacing
    spacing.set(_ns("beforeAutospacing"), "0")
    spacing.set(_ns("afterAutospacing"), "0")

    # Optional line spacing
    if line_spacing is not None:
        spacing.set(_ns("line"), str(int(line_spacing * 240)))
        spacing.set(_ns("lineRule"), line_rule)


def set_paragraph_alignment(para, alignment: str) -> None:
    """Write w:jc directly at OOXML level, ensuring it overrides style-based alignment.

    alignment must be one of: "left", "center", "right", "both", "distribute".
    Use this instead of python-docx para.alignment when style inheritance
    might otherwise override the setting.
    """
    p_elem = para._element
    pPr = p_elem.find(_ns("pPr"))
    if pPr is None:
        pPr = etree.SubElement(p_elem, _ns("pPr"))
    jc = pPr.find(_ns("jc"))
    if jc is None:
        jc = etree.SubElement(pPr, _ns("jc"))
    jc.set(_ns("val"), alignment)

from tvba_utils import cm_to_points


def _sync_numbering_with_titles_legacy(doc, settings, *, _paragraphs=None) -> None:
    """Update numbering definitions so auto-generated list numbers match title formatting.

    When paragraphs use Word multilevel lists for headings, the numbers are
    rendered by Word from numbering definitions in numbering.xml, not from
    paragraph runs. This function updates the w:rPr inside w:lvl elements to
    match the corresponding title settings.
    """
    from tvba_utils import size_label_to_points

    numbering_part = doc.part.numbering_part
    if numbering_part is None:
        return

    # Collect numIds used by heading paragraphs (have outline level)
    heading_num_ids = set()
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    paragraphs = _paragraphs if _paragraphs is not None else doc.paragraphs
    for para in paragraphs:
        pPr = para._element.find(f".//{{{W}}}pPr")
        if pPr is None:
            continue
        # Check if paragraph has outline level (title)
        outline = pPr.find(f"{{{W}}}outlineLvl")
        if outline is None:
            continue
        # Check if paragraph also has list numbering
        numPr = pPr.find(f"{{{W}}}numPr")
        if numPr is None:
            continue
        numId_elem = numPr.find(f"{{{W}}}numId")
        if numId_elem is not None:
            val = numId_elem.get(f"{{{W}}}val")
            if val is not None:
                heading_num_ids.add(val)

    if not heading_num_ids:
        return

    # Map numId -> abstractNumId
    num_id_to_abstract = {}
    for num in numbering_part._element.findall(f"{{{W}}}num"):
        num_id = num.get(f"{{{W}}}numId")
        abstract_id_elem = num.find(f"{{{W}}}abstractNumId")
        if abstract_id_elem is not None:
            abstract_id = abstract_id_elem.get(f"{{{W}}}val")
            if num_id is not None and abstract_id is not None:
                num_id_to_abstract[num_id] = abstract_id

    # Find abstractNums used by headings and update their levels
    abstract_ids_to_update = {
        num_id_to_abstract[nid]
        for nid in heading_num_ids
        if nid in num_id_to_abstract
    }

    for abstract_num in numbering_part._element.findall(f"{{{W}}}abstractNum"):
        abstract_id = abstract_num.get(f"{{{W}}}abstractNumId")
        if abstract_id not in abstract_ids_to_update:
            continue

        for lvl in abstract_num.findall(f"{{{W}}}lvl"):
            ilvl_str = lvl.get(f"{{{W}}}ilvl")
            if ilvl_str is None:
                continue
            ilvl = int(ilvl_str)
            if not (0 <= ilvl <= 4):
                continue

            title_settings = settings.titles[ilvl]
            half_points = str(int(size_label_to_points(title_settings.size) * 2))

            rPr = lvl.find(f"{{{W}}}rPr")
            if rPr is None:
                rPr = etree.SubElement(lvl, f"{{{W}}}rPr")

            # Font
            rFonts = rPr.find(f"{{{W}}}rFonts")
            if rFonts is None:
                rFonts = etree.SubElement(rPr, f"{{{W}}}rFonts")
            rFonts.set(f"{{{W}}}ascii", "Times New Roman")
            rFonts.set(f"{{{W}}}hAnsi", "Times New Roman")
            rFonts.set(f"{{{W}}}eastAsia", title_settings.font)

            # Size
            sz = rPr.find(f"{{{W}}}sz")
            if sz is None:
                sz = etree.SubElement(rPr, f"{{{W}}}sz")
            sz.set(f"{{{W}}}val", half_points)

            szCs = rPr.find(f"{{{W}}}szCs")
            if szCs is None:
                szCs = etree.SubElement(rPr, f"{{{W}}}szCs")
            szCs.set(f"{{{W}}}val", half_points)

            # Bold
            for tag in ("b", "bCs"):
                b = rPr.find(f"{{{W}}}{tag}")
                if b is None:
                    b = etree.SubElement(rPr, f"{{{W}}}{tag}")
                if title_settings.bold:
                    b.attrib.pop(f"{{{W}}}val", None)
                else:
                    b.set(f"{{{W}}}val", "0")

            pPr = lvl.find(f"{{{W}}}pPr")
            if pPr is None:
                pPr = etree.SubElement(lvl, f"{{{W}}}pPr")
            ind = pPr.find(f"{{{W}}}ind")
            if ind is None:
                ind = etree.SubElement(pPr, f"{{{W}}}ind")

            for attr in ("left", "right", "firstLine", "hanging", "firstLineChars", "hangingChars"):
                key = f"{{{W}}}{attr}"
                if key in ind.attrib:
                    del ind.attrib[key]
            ind.set(f"{{{W}}}leftChars", str(int(title_settings.left_indent_chars * 100)))
            ind.set(f"{{{W}}}rightChars", str(int(title_settings.right_indent_chars * 100)))
            if title_settings.special_indent == "首行缩进" and title_settings.special_indent_chars:
                ind.set(f"{{{W}}}firstLineChars", str(int(title_settings.special_indent_chars * 100)))
            elif title_settings.special_indent == "悬挂缩进" and title_settings.special_indent_chars:
                ind.set(f"{{{W}}}hangingChars", str(int(title_settings.special_indent_chars * 100)))
            else:
                ind.set(f"{{{W}}}firstLineChars", "0")
                ind.set(f"{{{W}}}firstLine", "0")


def sync_numbering_with_titles(doc, settings, *, _paragraphs=None) -> None:
    """Update list marker formatting for both title and body paragraphs.

    Word-generated list markers such as "1)" live in numbering.xml rather than
    in paragraph runs. Title markers follow title settings; ordinary list
    markers follow body settings so they are refreshed with the surrounding text.
    """
    from tvba_utils import size_label_to_points

    numbering_part = doc.part.numbering_part
    if numbering_part is None:
        return

    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    paragraphs = _paragraphs if _paragraphs is not None else doc.paragraphs
    used_levels = {}

    for para in paragraphs:
        pPr = para._element.find(f".//{{{W}}}pPr")
        if pPr is None:
            continue
        numPr = pPr.find(f"{{{W}}}numPr")
        if numPr is None:
            continue
        numId_elem = numPr.find(f"{{{W}}}numId")
        ilvl_elem = numPr.find(f"{{{W}}}ilvl")
        if numId_elem is None or ilvl_elem is None:
            continue
        num_id = numId_elem.get(f"{{{W}}}val")
        ilvl_str = ilvl_elem.get(f"{{{W}}}val")
        if num_id is None or ilvl_str is None:
            continue
        try:
            ilvl = int(ilvl_str)
        except ValueError:
            continue

        fmt = ("body", None)
        outline = pPr.find(f"{{{W}}}outlineLvl")
        if outline is not None:
            outline_val = outline.get(f"{{{W}}}val")
            try:
                outline_level = int(outline_val) if outline_val is not None else None
            except ValueError:
                outline_level = None
            if outline_level is not None and 0 <= outline_level <= 4:
                fmt = ("title", outline_level)

        key = (num_id, ilvl)
        if used_levels.get(key, ("body", None))[0] != "title":
            used_levels[key] = fmt

    if not used_levels:
        return

    num_id_to_abstract = {}
    for num in numbering_part._element.findall(f"{{{W}}}num"):
        num_id = num.get(f"{{{W}}}numId")
        abstract_id_elem = num.find(f"{{{W}}}abstractNumId")
        if abstract_id_elem is None:
            continue
        abstract_id = abstract_id_elem.get(f"{{{W}}}val")
        if num_id is not None and abstract_id is not None:
            num_id_to_abstract[num_id] = abstract_id

    abstract_level_formats = {}
    for (num_id, ilvl), fmt in used_levels.items():
        abstract_id = num_id_to_abstract.get(num_id)
        if abstract_id is not None:
            abstract_level_formats[(abstract_id, ilvl)] = fmt

    for abstract_num in numbering_part._element.findall(f"{{{W}}}abstractNum"):
        abstract_id = abstract_num.get(f"{{{W}}}abstractNumId")
        for lvl in abstract_num.findall(f"{{{W}}}lvl"):
            ilvl_str = lvl.get(f"{{{W}}}ilvl")
            if ilvl_str is None:
                continue
            try:
                ilvl = int(ilvl_str)
            except ValueError:
                continue
            fmt = abstract_level_formats.get((abstract_id, ilvl))
            if fmt is None:
                continue

            fmt_kind, fmt_level = fmt
            if fmt_kind == "title" and fmt_level is not None:
                title_settings = settings.titles[fmt_level]
                eastasia_font = title_settings.font
                size_pt = size_label_to_points(title_settings.size)
                bold = title_settings.bold
                left_chars = title_settings.left_indent_chars
                right_chars = title_settings.right_indent_chars
                special_indent = title_settings.special_indent
                special_indent_chars = title_settings.special_indent_chars
            else:
                eastasia_font = settings.body.font
                size_pt = size_label_to_points(settings.body.size)
                bold = False
                left_chars = 0.0
                right_chars = 0.0
                special_indent = "\u65e0"
                special_indent_chars = 0.0

            half_points = str(int(size_pt * 2))
            rPr = lvl.find(f"{{{W}}}rPr")
            if rPr is None:
                rPr = etree.SubElement(lvl, f"{{{W}}}rPr")

            rFonts = rPr.find(f"{{{W}}}rFonts")
            if rFonts is None:
                rFonts = etree.SubElement(rPr, f"{{{W}}}rFonts")
            rFonts.set(f"{{{W}}}ascii", "Times New Roman")
            rFonts.set(f"{{{W}}}hAnsi", "Times New Roman")
            rFonts.set(f"{{{W}}}eastAsia", eastasia_font)

            for tag in ("sz", "szCs"):
                elem = rPr.find(f"{{{W}}}{tag}")
                if elem is None:
                    elem = etree.SubElement(rPr, f"{{{W}}}{tag}")
                elem.set(f"{{{W}}}val", half_points)

            for tag in ("b", "bCs"):
                for old in list(rPr.findall(f"{{{W}}}{tag}")):
                    rPr.remove(old)
                elem = etree.SubElement(rPr, f"{{{W}}}{tag}")
                if not bold:
                    elem.set(f"{{{W}}}val", "0")

            pPr = lvl.find(f"{{{W}}}pPr")
            if pPr is None:
                pPr = etree.SubElement(lvl, f"{{{W}}}pPr")
            ind = pPr.find(f"{{{W}}}ind")
            if ind is None:
                ind = etree.SubElement(pPr, f"{{{W}}}ind")

            for attr in ("left", "right", "firstLine", "hanging", "firstLineChars", "hangingChars"):
                ind.attrib.pop(f"{{{W}}}{attr}", None)
            ind.set(f"{{{W}}}leftChars", str(int(left_chars * 100)))
            ind.set(f"{{{W}}}rightChars", str(int(right_chars * 100)))
            if special_indent == "\u9996\u884c\u7f29\u8fdb" and special_indent_chars:
                ind.set(f"{{{W}}}firstLineChars", str(int(special_indent_chars * 100)))
            elif special_indent == "\u60ac\u6302\u7f29\u8fdb" and special_indent_chars:
                ind.set(f"{{{W}}}hangingChars", str(int(special_indent_chars * 100)))
            else:
                ind.set(f"{{{W}}}firstLineChars", "0")
                ind.set(f"{{{W}}}firstLine", "0")


def set_table_layout(table, mode: str) -> None:
    """Set table column layout mode.

    mode:
      - "window"  — autofit to page width (w:tblLayout val="autofit", w:tblW 5000/auto)
      - "content" — autofit to cell content (w:tblLayout val="autofit", w:tblW auto)
      - "fixed"   — fixed column widths (w:tblLayout val="fixed")
    """
    tblPr = table._element.find(_ns("tblPr"))
    if tblPr is None:
        tblPr = etree.SubElement(table._element, _ns("tblPr"))

    # Layout algorithm
    layout = tblPr.find(_ns("tblLayout"))
    if layout is None:
        layout = etree.SubElement(tblPr, _ns("tblLayout"))
    layout.set(_ns("val"), "fixed" if mode == "fixed" else "autofit")

    # Table width
    tblW = tblPr.find(_ns("tblW"))
    if tblW is None:
        tblW = etree.SubElement(tblPr, _ns("tblW"))
    if mode == "window":
        tblW.set(_ns("w"), "5000")
        tblW.set(_ns("type"), "pct")
    elif mode == "content":
        tblW.set(_ns("w"), "0")
        tblW.set(_ns("type"), "auto")
    else:
        tblW.set(_ns("w"), "0")
        tblW.set(_ns("type"), "auto")


def set_table_layout_window(table) -> None:
    """Deprecated: use set_table_layout(table, 'window') instead."""
    set_table_layout(table, "window")


def set_table_layout_content(table) -> None:
    """Deprecated: use set_table_layout(table, 'fixed') instead."""
    set_table_layout(table, "fixed")


def set_table_borders(table, *, line_width_pt: float) -> None:
    """Set all table borders to a uniform width."""
    tblPr = table._element.find(_ns("tblPr"))
    if tblPr is None:
        tblPr = etree.SubElement(table._element, _ns("tblPr"))
    borders = tblPr.find(_ns("tblBorders"))
    if borders is None:
        borders = etree.SubElement(tblPr, _ns("tblBorders"))
    sz = str(max(2, int(round(line_width_pt * 8))))  # table borders use eighths of a point
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = borders.find(_ns(side))
        if border is None:
            border = etree.SubElement(borders, _ns(side))
        border.set(_ns("val"), "single")
        border.set(_ns("sz"), sz)
        border.set(_ns("space"), "0")
        border.set(_ns("color"), "auto")


def set_table_alignment(table, alignment: str = "center") -> None:
    """Set table horizontal alignment via w:tblPr/w:jc.

    alignment: "left", "center", "right"
    """
    tblPr = table._element.find(_ns("tblPr"))
    if tblPr is None:
        tblPr = etree.SubElement(table._element, _ns("tblPr"))
    jc = tblPr.find(_ns("jc"))
    if jc is None:
        jc = etree.SubElement(tblPr, _ns("jc"))
    jc.set(_ns("val"), alignment)


def set_snap_to_grid(para, enabled: bool = True) -> None:
    """Set w:snapToGrid on paragraph for grid alignment."""
    pPr = para._element.find(_ns("pPr"))
    if pPr is None:
        pPr = etree.SubElement(para._element, _ns("pPr"))
    snap = pPr.find(_ns("snapToGrid"))
    if enabled:
        if snap is None:
            snap = etree.SubElement(pPr, _ns("snapToGrid"))
        snap.set(_ns("val"), "true")
    else:
        if snap is not None:
            pPr.remove(snap)


def set_auto_space_de(para, enabled: bool = True) -> None:
    """Set w:autoSpaceDE (auto adjust right indent for East Asian text)."""
    pPr = para._element.find(_ns("pPr"))
    if pPr is None:
        pPr = etree.SubElement(para._element, _ns("pPr"))
    auto = pPr.find(_ns("autoSpaceDE"))
    if enabled:
        if auto is None:
            auto = etree.SubElement(pPr, _ns("autoSpaceDE"))
        auto.set(_ns("val"), "true")
    else:
        if auto is not None:
            pPr.remove(auto)


def set_row_height_at_least(row, height_cm: float) -> None:
    """Set row height to at least the given cm."""
    trPr = row._tr.find(_ns("trPr"))
    if trPr is None:
        trPr = etree.SubElement(row._tr, _ns("trPr"))
    trHeight = trPr.find(_ns("trHeight"))
    if trHeight is None:
        trHeight = etree.SubElement(trPr, _ns("trHeight"))
    points = cm_to_points(height_cm)
    twips = int(points * 20)
    trHeight.set(_ns("val"), str(twips))
    trHeight.set(_ns("hRule"), "atLeast")


def get_effective_run_fonts(run, doc=None) -> dict[str, str | None]:
    """Resolve effective fonts for a run by walking the style inheritance chain.

    Returns dict with keys: "ascii", "hAnsi", "eastAsia".
    Values are font names, "theme:<themeName>" if set via theme reference,
    or None if not resolvable at all.
    """
    fonts = {"ascii": None, "hAnsi": None, "eastAsia": None}
    font_attr_map = {"ascii": "ascii", "hAnsi": "hAnsi", "eastAsia": "eastAsia"}
    theme_attr_map = {"ascii": "asciiTheme", "hAnsi": "hAnsiTheme", "eastAsia": "eastAsiaTheme"}

    def _read_fonts_from_rfonts(rFonts):
        for key, attr in font_attr_map.items():
            if fonts[key] is None:
                val = rFonts.get(_ns(attr))
                if val:
                    fonts[key] = val
        for key, attr in theme_attr_map.items():
            if fonts[key] is None:
                val = rFonts.get(_ns(attr))
                if val:
                    fonts[key] = f"theme:{val}"

    # 1. Check run direct w:rPr/w:rFonts
    rPr = run._element.find(_ns("rPr"))
    if rPr is not None:
        rFonts = rPr.find(_ns("rFonts"))
        if rFonts is not None:
            _read_fonts_from_rfonts(rFonts)

    if all(fonts[k] for k in fonts):
        return fonts

    # 2. Check paragraph style run properties (and basedOn chain)
    try:
        para = run._parent  # python-docx paragraph
    except Exception:
        return fonts

    style = para.style if para is not None else None
    visited = set()
    while style is not None:
        style_id = id(style)
        if style_id in visited:
            break
        visited.add(style_id)

        style_rPr = style.element.find(_ns("rPr"))
        if style_rPr is not None:
            style_rFonts = style_rPr.find(_ns("rFonts"))
            if style_rFonts is not None:
                _read_fonts_from_rfonts(style_rFonts)

        if all(fonts[k] for k in fonts):
            return fonts

        # Follow basedOn chain
        basedOn = style.element.find(_ns("basedOn"))
        if basedOn is None:
            break
        based_on_val = basedOn.get(_ns("val"))
        if based_on_val is None or doc is None:
            break
        try:
            style = doc.styles[based_on_val]
        except KeyError:
            break

    # 3. Check document defaults
    if doc is not None:
        try:
            doc_defaults = doc.styles.element.find(_ns("docDefaults"))
            if doc_defaults is not None:
                rPrDefault = doc_defaults.find(_ns("rPrDefault"))
                if rPrDefault is not None:
                    rPr = rPrDefault.find(_ns("rPr"))
                    if rPr is not None:
                        rFonts = rPr.find(_ns("rFonts"))
                        if rFonts is not None:
                            _read_fonts_from_rfonts(rFonts)
        except Exception:
            pass

    return fonts

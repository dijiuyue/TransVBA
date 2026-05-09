"""OOXML helpers using lxml direct element manipulation.

Thin wrappers that look like python-docx API but operate at the lxml level
for attributes that python-docx does not expose.
"""
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


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

        # Set bold
        if bold:
            b = rPr.find(f"{{{W}}}b")
            if b is None:
                etree.SubElement(rPr, f"{{{W}}}b")
        else:
            b = rPr.find(f"{{{W}}}b")
            if b is not None:
                rPr.remove(b)


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


def get_effective_outline_level(paragraph, styles=None) -> int | None:
    """Return the effective outline level (0-8) for a paragraph.

    Checks the paragraph's direct w:pPr/w:outlineLvl first, then falls back
    to the paragraph style (and its basedOn chain). Returns None if no
    outline level is found.
    """
    # 1. Direct paragraph-level outline
    pPr = paragraph._element.find(_ns("pPr"))
    if pPr is not None:
        outline = pPr.find(_ns("outlineLvl"))
        if outline is not None:
            val = outline.get(_ns("val"))
            if val is not None:
                return int(val)

    # 2. Style-level outline (follow basedOn chain)
    style = paragraph.style
    while style is not None:
        style_pPr = style.element.find(_ns("pPr"))
        if style_pPr is not None:
            outline = style_pPr.find(_ns("outlineLvl"))
            if outline is not None:
                val = outline.get(_ns("val"))
                if val is not None:
                    return int(val)

        # Follow basedOn chain
        basedOn = style.element.find(_ns("basedOn"))
        if basedOn is None:
            break
        based_on_val = basedOn.get(_ns("val"))
        if based_on_val is None or styles is None:
            break
        try:
            style = styles[based_on_val]
        except KeyError:
            break

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
    # paragraph_format._element is the w:p element; we need w:pPr
    p_elem = paragraph_format._element
    pPr = p_elem.find(_ns("pPr"))
    if pPr is None:
        pPr = etree.SubElement(p_elem, _ns("pPr"))
    ind = pPr.find(_ns("ind"))
    if ind is None:
        ind = etree.SubElement(pPr, _ns("ind"))

    twips_per_char = 240  # 12 pt * 20 twips/pt

    if left_chars:
        ind.set(_ns("left"), str(int(left_chars * twips_per_char)))
    elif _ns("left") in ind.attrib:
        del ind.attrib[_ns("left")]

    if right_chars:
        ind.set(_ns("right"), str(int(right_chars * twips_per_char)))
    elif _ns("right") in ind.attrib:
        del ind.attrib[_ns("right")]

    # Clear existing special indent attrs
    for attr in (_ns("firstLine"), _ns("hanging")):
        if attr in ind.attrib:
            del ind.attrib[attr]

    if special_kind == "首行缩进" and special_chars:
        ind.set(_ns("firstLine"), str(int(special_chars * twips_per_char)))
    elif special_kind == "悬挂缩进" and special_chars:
        ind.set(_ns("hanging"), str(int(special_chars * twips_per_char)))


def set_before_after_lines(paragraph_format, *, before_lines: float, after_lines: float) -> None:
    """Set paragraph spacing in line units (w:beforeLines / w:afterLines).

    1 line = 100 hundredths of a line (same as VBA LineUnitBefore).
    """
    # paragraph_format._element is the w:p element; we need w:pPr
    p_elem = paragraph_format._element
    pPr = p_elem.find(_ns("pPr"))
    if pPr is None:
        pPr = etree.SubElement(p_elem, _ns("pPr"))
    spacing = pPr.find(_ns("spacing"))
    if spacing is None:
        spacing = etree.SubElement(pPr, _ns("spacing"))
    spacing.set(_ns("beforeLines"), str(int(before_lines * 100)))
    spacing.set(_ns("afterLines"), str(int(after_lines * 100)))


from tvba_utils import cm_to_points


def sync_numbering_with_titles(doc, settings) -> None:
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
    for para in doc.paragraphs:
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
            if title_settings.bold:
                b = rPr.find(f"{{{W}}}b")
                if b is None:
                    etree.SubElement(rPr, f"{{{W}}}b")
            else:
                b = rPr.find(f"{{{W}}}b")
                if b is not None:
                    rPr.remove(b)


def set_table_layout_window(table) -> None:
    """Set table to autofit to window (AutoFitBehavior=2)."""
    tblPr = table._element.find(_ns("tblPr"))
    if tblPr is None:
        tblPr = etree.SubElement(table._element, _ns("tblPr"))
    layout = tblPr.find(_ns("tblLayout"))
    if layout is None:
        layout = etree.SubElement(tblPr, _ns("tblLayout"))
    layout.set(_ns("val"), "autofit")


def set_table_layout_content(table) -> None:
    """Set table to autofit to content (AutoFitBehavior=1)."""
    tblPr = table._element.find(_ns("tblPr"))
    if tblPr is None:
        tblPr = etree.SubElement(table._element, _ns("tblPr"))
    layout = tblPr.find(_ns("tblLayout"))
    if layout is None:
        layout = etree.SubElement(tblPr, _ns("tblLayout"))
    layout.set(_ns("val"), "fixed")


def set_table_borders(table, *, line_width_pt: float) -> None:
    """Set all table borders to a uniform width."""
    tblPr = table._element.find(_ns("tblPr"))
    if tblPr is None:
        tblPr = etree.SubElement(table._element, _ns("tblPr"))
    borders = tblPr.find(_ns("tblBorders"))
    if borders is None:
        borders = etree.SubElement(tblPr, _ns("tblBorders"))
    sz = str(int(line_width_pt * 20))  # half-points (1 pt = 20 half-points)
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = borders.find(_ns(side))
        if border is None:
            border = etree.SubElement(borders, _ns(side))
        border.set(_ns("val"), "single")
        border.set(_ns("sz"), sz)
        border.set(_ns("space"), "0")
        border.set(_ns("color"), "auto")


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

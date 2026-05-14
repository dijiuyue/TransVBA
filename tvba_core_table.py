"""Table + table caption formatting.

Corresponds to VBA FormatModule.bas:
  - RefreshTableFormat
  - SetTableTitle
  - FindTableCaptionRange
  - FindCaptionInShapes
  - IsTableCaptionLine
"""
from tvba_core_oox import (
    set_table_layout,
    set_table_borders,
    set_table_alignment,
    set_row_height_at_least,
    apply_indent_chars,
    set_before_after_lines,
    format_all_runs_in_paragraph,
)
import re

from tvba_utils import clean_para_text, size_label_to_points, cm_to_points

# VBA pattern: ^表\s*\d+(\.\d+)*-\d+[\t ]+.+$
# Use \s+ to match any whitespace (space, tab, nbsp, fullwidth space, etc.)
_TABLE_CAPTION_RE = re.compile(
    r"^(?:表|table)\s*\d+(?:\.\d+)*-\d+\s+.+$",
    re.IGNORECASE,
)


def is_table_caption_line(text: str) -> bool:
    """Check if text is a table caption."""
    text = clean_para_text(text)
    return bool(_TABLE_CAPTION_RE.match(text))


def find_table_caption(table, doc, max_up_paragraphs: int = 10, *, _tables=None, _paragraphs=None):
    """Find the caption paragraph preceding a table."""
    # Find table index in document
    tables = _tables if _tables is not None else doc.tables
    table_index = None
    for i, t in enumerate(tables):
        if t._element is table._element:
            table_index = i
            break

    if table_index is None:
        return None

    # Count paragraphs before this table
    paragraphs_before = 0
    for element in doc.element.body:
        if element is table._element:
            break
        if element.tag.endswith("}p"):
            paragraphs_before += 1

    # Search backwards up to max_up_paragraphs
    paragraphs = _paragraphs if _paragraphs is not None else doc.paragraphs
    for i in range(1, max_up_paragraphs + 1):
        idx = paragraphs_before - i
        if idx < 0:
            break
        para = paragraphs[idx]
        if is_table_caption_line(para.text):
            return para

    # Search in shapes/text frames (WPS templates often put captions in text boxes)
    # python-docx cannot modify shape text, but we can detect it for awareness
    try:
        from lxml import etree
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        # Search all w:drawing elements in the document for text box content
        for drawing in doc.element.body.findall(f".//{{{W}}}drawing"):
            # Look for text body (a:txBody) within the drawing
            tx_body = drawing.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}txBody")
            if tx_body is not None:
                # Extract all paragraph text from the text body
                texts = []
                for ap in tx_body.findall("{http://schemas.openxmlformats.org/drawingml/2006/main}p"):
                    para_text = ""
                    for r in ap.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}t"):
                        if r.text:
                            para_text += r.text
                    if is_table_caption_line(para_text):
                        # Found caption in shape — cannot format via python-docx,
                        # but we can at least acknowledge it exists
                        return None  # Return None since we can't return a paragraph object
    except Exception:
        pass

    return None


def apply_table_caption(para, settings) -> None:
    """Apply formatting to a table caption paragraph."""
    # Normalize space between number and caption text: "表 1-1  测试" → "表 1-1 测试"
    _normalize_caption_space(para)
    format_all_runs_in_paragraph(
        para,
        ascii_font="Times New Roman",
        eastasia_font=settings.title_font,
        size_pt=size_label_to_points(settings.title_size),
        bold=settings.title_bold,
    )

    para.alignment = 1  # Center alignment for captions

    # Clear any inherited indentation (body first-line/hanging indent)
    apply_indent_chars(
        para.paragraph_format,
        left_chars=0.0,
        right_chars=0.0,
        special_kind="无",
        special_chars=0.0,
    )

    set_before_after_lines(
        para.paragraph_format,
        before_lines=0.0,
        after_lines=0.0,
    )

    # Line spacing
    pPr = para._element.find(".//w:pPr", {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
    if pPr is not None:
        from lxml import etree
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        spacing = pPr.find(f"{{{W}}}spacing")
        if spacing is None:
            spacing = etree.SubElement(pPr, f"{{{W}}}spacing")
        spacing.set(f"{{{W}}}line", str(int(settings.title_spacing * 240)))
        spacing.set(f"{{{W}}}lineRule", "auto")


def apply_table_body(table, settings) -> None:
    """Apply formatting to table body."""
    # Table centering
    set_table_alignment(table, "center")

    # Column layout mode
    set_table_layout(table, settings.auto_fit_mode)

    # Borders
    set_table_borders(table, line_width_pt=settings.line_width_pt)

    # Row height
    for row in table.rows:
        set_row_height_at_least(row, settings.row_height_cm)

    # Cell font — first row (header) bold, rest normal
    for row_idx, row in enumerate(table.rows):
        bold_row = (row_idx == 0)
        row_spacing = 1.0 if bold_row else settings.spacing  # 单倍 for header row
        for cell in row.cells:
            for para in cell.paragraphs:
                format_all_runs_in_paragraph(
                    para,
                    ascii_font="Times New Roman",
                    eastasia_font=settings.body_font,
                    size_pt=size_label_to_points(settings.body_size),
                    bold=bold_row,
                )
                # Line spacing
                pPr = para._element.find(".//w:pPr", {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
                if pPr is not None:
                    from lxml import etree
                    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    spacing = pPr.find(f"{{{W}}}spacing")
                    if spacing is None:
                        spacing = etree.SubElement(pPr, f"{{{W}}}spacing")
                    spacing.set(f"{{{W}}}line", str(int(row_spacing * 240)))
                    spacing.set(f"{{{W}}}lineRule", "auto")


def _normalize_caption_space(para) -> None:
    """Ensure exactly one space between caption number and title text.

    "表 1-1    测试" or "表 1-1\t测试" → "表 1-1 测试"

    Only normalizes when the full caption is within a single run,
    to avoid corrupting multi-run paragraphs.
    """
    import re
    for run in para.runs:
        m = re.match(r'^([表图]\s*\d+(?:\.\d+)*-\d+)(\s{2,})(.+)$', run.text)
        if m:
            run.text = m.group(1) + ' ' + m.group(3)
            return


def refresh_all(doc, settings, *, _paragraphs=None, _tables=None) -> None:
    """Refresh all tables and their captions."""
    tables = _tables if _tables is not None else doc.tables
    for table in tables:
        caption = find_table_caption(table, doc, _tables=tables, _paragraphs=_paragraphs)
        if caption is not None:
            apply_table_caption(caption, settings)
        apply_table_body(table, settings)

"""ASCII font normalization and text fixes.

Corresponds to VBA FormatModule.bas:
  - NormalizeAsciiFont
  - ApplyBrackets
  - AddPeriodIfNeeded
  - SyncNumberFontWithBody
"""
import re

from tvba_core_oox import set_ascii_font

_ASCII_RE = re.compile(r"[\x00-\x7F]+")


def unify_ascii_font(doc, font_name: str = "Times New Roman") -> None:
    """Set all ASCII-only runs to the specified font."""
    for para in doc.paragraphs:
        for run in para.runs:
            text = run.text
            if text and all(ord(c) < 128 for c in text):
                set_ascii_font(run, font_name)
        # Also handle table cells
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        text = run.text
                        if text and all(ord(c) < 128 for c in text):
                            set_ascii_font(run, font_name)


_BRACKET_RE = re.compile(r"[（(].*?[）)]")


def apply_brackets(para, text: str) -> None:
    """Make text within brackets bold.

    Finds the first occurrence of bracketed text using fullwidth
    （…）or halfwidth (…) brackets and applies bold formatting.
    Only processes the first occurrence per paragraph (VBA behavior).
    """
    m = _BRACKET_RE.search(text)
    if not m:
        return

    bracket_start = m.start()
    bracket_end = m.end()

    # Walk runs to find which run(s) contain the bracketed text
    pos = 0
    for run in para.runs:
        run_text = run.text
        run_len = len(run_text)
        run_start = pos
        run_end = pos + run_len

        # Check if this run overlaps with the bracketed region
        if run_end > bracket_start and run_start < bracket_end:
            # Determine the overlap range within this run
            overlap_start = max(bracket_start, run_start) - run_start
            overlap_end = min(bracket_end, run_end) - run_start

            if overlap_start == 0 and overlap_end == run_len:
                # Entire run is bracketed
                run.font.bold = True
            else:
                # Partial overlap — split the run
                before = run_text[:overlap_start]
                bracketed = run_text[overlap_start:overlap_end]
                after = run_text[overlap_end:]

                # Replace run text
                run.text = before + bracketed + after
                # Re-find the run that now contains bracketed text
                # Setting text may have created a new run element;
                # set bold on the current run's font for the bracketed portion
                # Since python-docx run.text setter preserves the run,
                # we need to split into separate runs
                run.text = before
                new_run = para._element.find(".//w:r", {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
                # Use insert approach: add new runs after current one
                from lxml import etree
                W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

                # Clear current run text to 'before'
                run.text = before

                # Create bold run for bracketed text
                r_elem = run._element
                new_r = etree.Element(f"{{{W}}}r")
                rPr = etree.SubElement(new_r, f"{{{W}}}rPr")
                b = etree.SubElement(rPr, f"{{{W}}}b")
                t = etree.SubElement(new_r, f"{{{W}}}t")
                t.text = bracketed
                r_elem.addnext(new_r)

                # Create normal run for after text
                if after:
                    after_r = etree.Element(f"{{{W}}}r")
                    t_after = etree.SubElement(after_r, f"{{{W}}}t")
                    t_after.text = after
                    new_r.addnext(after_r)

                # Only process first occurrence
                return

        pos += run_len
        if pos >= bracket_end:
            break


def add_period_if_needed(para) -> None:
    """Add Chinese period to paragraph if it doesn't end with punctuation.

    Skips empty paragraphs and TOC lines.
    """
    from tvba_core_toc import is_toc_paragraph

    if is_toc_paragraph(para):
        return

    text = para.text
    if not text or not text.strip():
        return

    stripped = text.rstrip()
    if not stripped:
        return

    # Check if already ends with a punctuation mark
    if stripped[-1] in _END_PUNCTUATION:
        return

    # Append Chinese period to the last run, or create a new run
    if para.runs:
        last_run = para.runs[-1]
        last_run.text = last_run.text.rstrip() + "。"
    else:
        # No runs — add a new run
        para.add_run("。")


_END_PUNCTUATION = frozenset("。.！!?？")


def sync_number_font_with_body(para) -> None:
    """Sync auto-number font with body font for list paragraphs.

    For paragraphs with auto-numbering, ensures the numbering font
    matches the body font. Since full COM list template access is
    limited in python-docx, this sets the font on the paragraph's
    runs as a basic fallback.
    """
    # Check if paragraph has list numbering
    pPr = para._element.find(".//w:pPr", {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
    if pPr is None:
        return

    # Check for numbering properties
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    numPr = pPr.find(f"{{{W}}}numPr")
    if numPr is None:
        return

    # Get the body font from the first run, or use a default
    body_font = None
    for run in para.runs:
        if run.font.name:
            body_font = run.font.name
            break

    if body_font is None:
        body_font = "Times New Roman"

    # Apply body font to all runs in the paragraph
    for run in para.runs:
        if run.font.name is None or run.font.name == "":
            run.font.name = body_font

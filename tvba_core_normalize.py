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


def apply_brackets(para, text: str) -> None:
    """Apply bracket normalization (placeholder for VBA ApplyBrackets)."""
    # VBA behavior: normalize fullwidth brackets to halfwidth
    # Implementation deferred to when specific test cases are identified
    pass


def add_period_if_needed(para) -> None:
    """Add period to title if missing (placeholder for VBA AddPeriodIfNeeded)."""
    # VBA behavior: add Chinese period to titles that don't end with punctuation
    pass


def sync_number_font_with_body(para) -> None:
    """Sync number font with body font (placeholder for VBA SyncNumberFontWithBody)."""
    # VBA behavior: ensure numbers in paragraph use body font
    pass

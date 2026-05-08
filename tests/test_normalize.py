import pytest
from docx import Document

from tvba_core_normalize import (
    unify_ascii_font,
    apply_brackets,
    add_period_if_needed,
    sync_number_font_with_body,
)


class TestUnifyAsciiFont:
    def test_sets_ascii_runs_to_times_new_roman(self):
        doc = Document()
        para = doc.add_paragraph("Hello World 123")
        unify_ascii_font(doc, "Times New Roman")
        for para in doc.paragraphs:
            for run in para.runs:
                assert run.font.name == "Times New Roman"

    def test_skips_cjk_characters(self):
        doc = Document()
        para = doc.add_paragraph("中文")
        # Should not crash
        unify_ascii_font(doc, "Times New Roman")


class TestApplyBrackets:
    def test_no_op_for_now(self):
        doc = Document()
        para = doc.add_paragraph("(test)")
        apply_brackets(para, "(test)")
        # Placeholder: function exists and doesn't crash


class TestAddPeriodIfNeeded:
    def test_adds_period_to_title(self):
        doc = Document()
        para = doc.add_paragraph("1 标题")
        add_period_if_needed(para)
        # Placeholder behavior


class TestSyncNumberFontWithBody:
    def test_syncs_number_font(self):
        doc = Document()
        para = doc.add_paragraph("123")
        sync_number_font_with_body(para)
        # Placeholder: function exists

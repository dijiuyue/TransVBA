import pytest
from docx import Document

from tvba_core_table import (
    is_table_caption_line,
    find_table_caption,
    apply_table_caption,
    apply_table_body,
    refresh_all,
)
from tvba_settings import TableSettings

class TestIsTableCaptionLine:
    def test_starts_with_biao(self):
        assert is_table_caption_line("表 1.1-1 示例表格") is True

    def test_starts_with_table(self):
        assert is_table_caption_line("Table 1 Example") is True

    def test_no_prefix(self):
        assert is_table_caption_line("示例表格") is False

    def test_case_insensitive(self):
        assert is_table_caption_line("TABLE 1 Example") is True

class TestApplyTableCaption:
    def test_applies_font_and_bold(self):
        doc = Document()
        para = doc.add_paragraph("表 1.1-1 示例")
        settings = TableSettings(title_font="黑体", title_size="小四", title_bold=True)
        apply_table_caption(para, settings)
        run = para.runs[0]
        assert run.font.bold is True

class TestApplyTableBody:
    def test_sets_borders(self):
        doc = Document()
        table = doc.add_table(rows=2, cols=2)
        settings = TableSettings(line_width_pt=1.0, auto_fit_window=True)
        apply_table_body(table, settings)
        tblPr = table._element.find(".//w:tblPr", {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
        borders = tblPr.find("w:tblBorders", {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
        assert borders is not None

class TestRefreshAll:
    def test_finds_and_formats_table(self):
        doc = Document()
        doc.add_paragraph("表 1.1-1 测试表格")
        table = doc.add_table(rows=2, cols=2)
        settings = TableSettings()
        refresh_all(doc, settings)
        # Caption should be formatted
        para = doc.paragraphs[0]
        assert para.runs[0].font.bold is True

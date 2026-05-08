import pytest
from docx import Document

from tvba_core_figure import (
    is_figure_caption_line,
    apply_figure_caption,
    refresh_all,
)
from tvba_settings import FigureSettings


class TestIsFigureCaptionLine:
    def test_starts_with_figure(self):
        assert is_figure_caption_line("图 1-1 示例图片") is True

    def test_starts_with_fig(self):
        assert is_figure_caption_line("Fig 1 Example") is True

    def test_no_prefix(self):
        assert is_figure_caption_line("示例图片") is False

    def test_case_insensitive(self):
        assert is_figure_caption_line("FIGURE 1 Example") is True


class TestApplyFigureCaption:
    def test_applies_font_and_bold(self):
        doc = Document()
        para = doc.add_paragraph("图 1-1 示例")
        settings = FigureSettings(title_font="黑体", title_size="小四", title_bold=True)
        apply_figure_caption(para, settings)
        run = para.runs[0]
        assert run.font.bold is True


class TestRefreshAll:
    def test_finds_and_formats_captions(self):
        doc = Document()
        doc.add_paragraph("图 1-1 测试图片")
        doc.add_paragraph("正文段落")
        doc.add_paragraph("图 2-1 另一个图片")
        settings = FigureSettings()
        refresh_all(doc, settings)
        assert doc.paragraphs[0].runs[0].font.bold is True
        assert doc.paragraphs[1].runs[0].font.bold is not True

import pytest
from docx import Document

from tvba_core_title import (
    identify_numeric_title_level,
    identify_level_from_number,
    normalize_number_string,
    apply_title_style,
    auto_detect_and_format,
)
from tvba_settings import TitleLevelSettings, BodySettings, FormatSettings


class TestNormalizeNumberString:
    def test_fullwidth_dot_to_halfwidth(self):
        assert normalize_number_string("1．2．3") == "1.2.3"

    def test_removes_trailing_dot(self):
        assert normalize_number_string("1.2.") == "1.2"

    def test_strips_whitespace(self):
        assert normalize_number_string("  1.1  ") == "1.1"

    def test_multiple_fullwidth_dots(self):
        assert normalize_number_string("１．２．３．４") == "1.2.3.4"


class TestIdentifyLevelFromNumber:
    def test_level_1_no_dot(self):
        assert identify_level_from_number("1") == 1

    def test_level_1_with_dot_zero(self):
        assert identify_level_from_number("1.0") == 1

    def test_level_2_one_dot(self):
        assert identify_level_from_number("1.1") == 2

    def test_level_3_two_dots(self):
        assert identify_level_from_number("1.1.2") == 3

    def test_level_4_three_dots(self):
        assert identify_level_from_number("1.1.2.3") == 4

    def test_level_5_four_dots(self):
        assert identify_level_from_number("1.1.2.3.4") == 5

    def test_level_0_too_many_dots(self):
        assert identify_level_from_number("1.1.2.3.4.5") == 0

    def test_level_0_empty(self):
        assert identify_level_from_number("") == 0


class TestIdentifyNumericTitleLevel:
    def test_level_1_simple(self):
        assert identify_numeric_title_level("1 引言") == 1

    def test_level_2_simple(self):
        assert identify_numeric_title_level("1.1 背景") == 2

    def test_level_3_simple(self):
        assert identify_numeric_title_level("1.1.1 详细背景") == 3

    def test_no_number_returns_0(self):
        assert identify_numeric_title_level("引言") == 0

    def test_number_without_space_returns_0(self):
        assert identify_numeric_title_level("1引言") == 0

    def test_fullwidth_dots_work(self):
        assert identify_numeric_title_level("1．1 背景") == 2

    def test_tab_separator_works(self):
        assert identify_numeric_title_level("1.1\t背景") == 2

    def test_trailing_dot_normalized(self):
        assert identify_numeric_title_level("1. 引言") == 1

    def test_level_1_requires_space_or_tab(self):
        assert identify_numeric_title_level("1引言") == 0
        assert identify_numeric_title_level("1 引言") == 1

    def test_too_many_dots_returns_0(self):
        assert identify_numeric_title_level("1.1.2.3.4.5 太深") == 0


class TestApplyTitleStyle:
    def test_applies_outline_level(self):
        doc = Document()
        para = doc.add_paragraph("1 标题")
        settings = TitleLevelSettings(font="黑体", size="三号", bold=True)
        body = BodySettings()
        apply_title_style(para, 1, settings, body)
        pPr = para._element.find(".//w:pPr", {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
        outline = pPr.find("w:outlineLvl", {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
        assert outline is not None
        assert outline.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") == "0"

    def test_applies_font_and_bold(self):
        doc = Document()
        para = doc.add_paragraph("1 标题")
        settings = TitleLevelSettings(font="黑体", size="三号", bold=True)
        body = BodySettings()
        apply_title_style(para, 1, settings, body)
        run = para.runs[0]
        assert run.font.bold is True

    def test_applies_center_alignment(self):
        doc = Document()
        para = doc.add_paragraph("1 标题")
        settings = TitleLevelSettings(alignment="居中")
        body = BodySettings()
        apply_title_style(para, 1, settings, body)
        assert para.alignment == 1


class TestAutoDetectAndFormat:
    def test_detects_and_formats_titles(self):
        doc = Document()
        doc.add_paragraph("1 一级标题")
        doc.add_paragraph("1.1 二级标题")
        doc.add_paragraph("正文段落")
        settings = FormatSettings()
        auto_detect_and_format(doc, settings)

        # Check first paragraph has outline level 0 (level 1)
        p1 = doc.paragraphs[0]
        pPr = p1._element.find(".//w:pPr", {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
        outline = pPr.find("w:outlineLvl", {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
        assert outline.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") == "0"

        # Check second paragraph has outline level 1 (level 2)
        p2 = doc.paragraphs[1]
        pPr2 = p2._element.find(".//w:pPr", {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
        outline2 = pPr2.find("w:outlineLvl", {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
        assert outline2.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") == "1"

        # Check body paragraph has no outline level
        p3 = doc.paragraphs[2]
        pPr3 = p3._element.find(".//w:pPr", {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
        if pPr3 is not None:
            outline3 = pPr3.find("w:outlineLvl", {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
            assert outline3 is None
        # If pPr3 is None, body paragraph has no formatting at all — also correct

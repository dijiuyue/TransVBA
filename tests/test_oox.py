import tempfile
from pathlib import Path
import pytest
from docx import Document
from lxml import etree

from tvba_core_oox import (
    set_far_east_font,
    set_ascii_font,
    set_outline_level,
    apply_indent_chars,
    set_before_after_lines,
)

NSMAP = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

class TestSetFarEastFont:
    def test_sets_rFonts_eastAsia_attribute(self):
        doc = Document()
        para = doc.add_paragraph("测试")
        run = para.runs[0]
        set_far_east_font(run, "黑体")
        rPr = run._element.find(".//w:rPr", NSMAP)
        assert rPr is not None
        rFonts = rPr.find("w:rFonts", NSMAP)
        assert rFonts is not None
        assert rFonts.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia") == "黑体"

class TestSetAsciiFont:
    def test_sets_font_name(self):
        doc = Document()
        para = doc.add_paragraph("Hello")
        run = para.runs[0]
        set_ascii_font(run, "Times New Roman")
        assert run.font.name == "Times New Roman"

class TestSetOutlineLevel:
    def test_sets_outline_level_zero(self):
        doc = Document()
        para = doc.add_paragraph("Title")
        set_outline_level(para, 0)
        pPr = para._element.find(".//w:pPr", NSMAP)
        assert pPr is not None
        outline = pPr.find("w:outlineLvl", NSMAP)
        assert outline is not None
        assert outline.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") == "0"

    def test_sets_outline_level_four(self):
        doc = Document()
        para = doc.add_paragraph("Title")
        set_outline_level(para, 4)
        pPr = para._element.find(".//w:pPr", NSMAP)
        outline = pPr.find("w:outlineLvl", NSMAP)
        assert outline.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") == "4"

class TestApplyIndentChars:
    def test_applies_left_indent_in_twips(self):
        doc = Document()
        para = doc.add_paragraph("Text")
        apply_indent_chars(
            para.paragraph_format,
            left_chars=2.0,
            right_chars=0.0,
            special_kind="无",
            special_chars=0.0,
        )
        pPr = para._element.find(".//w:pPr", NSMAP)
        ind = pPr.find("w:ind", NSMAP)
        assert ind is not None
        # 2 chars * 12 points/char * 20 twips/point = 480 twips
        assert ind.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}left") == "480"

    def test_applies_first_line_indent(self):
        doc = Document()
        para = doc.add_paragraph("Text")
        apply_indent_chars(
            para.paragraph_format,
            left_chars=0.0,
            right_chars=0.0,
            special_kind="首行缩进",
            special_chars=2.0,
        )
        pPr = para._element.find(".//w:pPr", NSMAP)
        ind = pPr.find("w:ind", NSMAP)
        assert ind.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}firstLine") == "480"

    def test_applies_hanging_indent(self):
        doc = Document()
        para = doc.add_paragraph("Text")
        apply_indent_chars(
            para.paragraph_format,
            left_chars=0.0,
            right_chars=0.0,
            special_kind="悬挂缩进",
            special_chars=2.0,
        )
        pPr = para._element.find(".//w:pPr", NSMAP)
        ind = pPr.find("w:ind", NSMAP)
        assert ind.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hanging") == "480"

class TestSetBeforeAfterLines:
    def test_sets_beforeLines_and_afterLines(self):
        doc = Document()
        para = doc.add_paragraph("Text")
        set_before_after_lines(
            para.paragraph_format,
            before_lines=0.5,
            after_lines=0.5,
        )
        pPr = para._element.find(".//w:pPr", NSMAP)
        spacing = pPr.find("w:spacing", NSMAP)
        assert spacing is not None
        assert spacing.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}beforeLines") == "50"
        assert spacing.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}afterLines") == "50"

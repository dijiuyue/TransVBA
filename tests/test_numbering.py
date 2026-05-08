import pytest
from docx import Document

from tvba_core_numbering import DocxListResolver, auto_select


class TestDocxListResolver:
    def test_no_numbering_returns_none(self):
        doc = Document()
        para = doc.add_paragraph("普通段落")
        resolver = DocxListResolver(doc)
        assert resolver.get_list_level(para) is None
        assert resolver.get_list_text(para) is None

    def test_returns_level_from_numPr(self):
        doc = Document()
        # Add a paragraph with numbering via OOXML
        para = doc.add_paragraph("列表项")
        pPr = para._element.get_or_add_pPr()
        from lxml import etree

        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        numPr = etree.SubElement(pPr, f"{{{W}}}numPr")
        ilvl = etree.SubElement(numPr, f"{{{W}}}ilvl")
        ilvl.set(f"{{{W}}}val", "2")
        numId = etree.SubElement(numPr, f"{{{W}}}numId")
        numId.set(f"{{{W}}}val", "1")

        resolver = DocxListResolver(doc)
        # Docx resolver returns ilvl + 1 as level
        assert resolver.get_list_level(para) == 3


class TestAutoSelect:
    def test_returns_resolver(self):
        resolver = auto_select(prefer_com=False)
        assert resolver is not None

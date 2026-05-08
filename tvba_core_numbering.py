"""Multi-level list resolver (COM bridge + docx fallback).

Corresponds to VBA FormatModule.bas:
  - IsMultiLevelListParagraph
  - ReportAllMultiLevelListLevels
"""
from typing import Protocol, runtime_checkable
from dataclasses import dataclass


@runtime_checkable
class ListResolver(Protocol):
    def get_list_level(self, para) -> int | None:
        """Return list level 1-9, or None if not a list paragraph."""
        ...

    def get_list_text(self, para) -> str | None:
        """Return rendered list text like '1.2.3', or None."""
        ...

    def diagnose(self, doc) -> list:
        """Return diagnostic entries for all list paragraphs."""
        ...


@dataclass
class DiagnosticEntry:
    text: str
    level: int | None
    list_text: str | None


class DocxListResolver:
    """Pure python-docx fallback: reads numPr/ilvl, simulates counting.

    Cannot reliably compute rendered list text, so returns None for get_list_text.
    """

    def __init__(self, doc):
        self.doc = doc

    def get_list_level(self, para) -> int | None:
        pPr = para._element.find(
            ".//w:pPr", {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        )
        if pPr is None:
            return None
        numPr = pPr.find(
            "w:numPr", {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        )
        if numPr is None:
            return None
        ilvl = numPr.find(
            "w:ilvl", {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        )
        if ilvl is not None:
            val = ilvl.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")
            if val is not None:
                return int(val) + 1  # Convert 0-indexed to 1-indexed
        return None

    def get_list_text(self, para) -> str | None:
        # Cannot reliably compute without parsing numbering definitions
        return None

    def diagnose(self, doc) -> list[DiagnosticEntry]:
        entries = []
        for para in doc.paragraphs:
            level = self.get_list_level(para)
            if level is not None:
                entries.append(
                    DiagnosticEntry(
                        text=para.text[:50],
                        level=level,
                        list_text=None,
                    )
                )
        return entries


class ComListResolver:
    """COM-based resolver using pywin32 Word automation.

    Provides 100% VBA-compatible ListLevelNumber and ListString.
    """

    def __init__(self, docx_path: str):
        self.docx_path = docx_path
        self.word = None
        self.doc = None

    def __enter__(self):
        import win32com.client

        self.word = win32com.client.Dispatch("Word.Application")
        self.word.Visible = False
        self.doc = self.word.Documents.Open(self.docx_path)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        try:
            if self.doc:
                self.doc.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if self.word:
                self.word.Quit()
        except Exception:
            pass
        return False

    def _get_com_paragraph(self, para):
        """Map a python-docx paragraph to the corresponding Word COM paragraph by index."""
        # Find the index of the paragraph in the python-docx document
        parent = para._element.getparent()
        if parent is None:
            return None
        try:
            idx = list(parent).index(para._element)
        except ValueError:
            return None
        # COM is 1-indexed
        return self.doc.Paragraphs(idx + 1)

    def get_list_level(self, para) -> int | None:
        com_para = self._get_com_paragraph(para)
        if com_para is None:
            return None
        try:
            list_format = com_para.Range.ListFormat
            if list_format.ListType == 0:  # wdListNoNumbering = 0
                return None
            level = list_format.ListLevelNumber
            return level if level >= 1 else None
        except Exception:
            return None

    def get_list_text(self, para) -> str | None:
        com_para = self._get_com_paragraph(para)
        if com_para is None:
            return None
        try:
            list_format = com_para.Range.ListFormat
            if list_format.ListType == 0:  # wdListNoNumbering = 0
                return None
            return list_format.ListString
        except Exception:
            return None

    def diagnose(self, doc) -> list[DiagnosticEntry]:
        entries = []
        for para in doc.paragraphs:
            level = self.get_list_level(para)
            if level is not None:
                entries.append(
                    DiagnosticEntry(
                        text=para.text[:50],
                        level=level,
                        list_text=self.get_list_text(para),
                    )
                )
        return entries


def auto_select(prefer_com: bool = True, docx_path: str | None = None) -> ListResolver:
    """Auto-select best available list resolver.

    If prefer_com is True and Word is available and docx_path is provided,
    returns a COM-based resolver. Otherwise returns DocxListResolver.
    """
    if prefer_com and docx_path is not None:
        try:
            import win32com.client

            word = win32com.client.Dispatch("Word.Application")
            word.Quit()
            return ComListResolver(docx_path)
        except Exception:
            pass
    return DocxListResolver(None)

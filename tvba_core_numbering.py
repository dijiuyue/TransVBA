"""Multi-level list resolver (COM bridge + docx fallback).

Corresponds to VBA FormatModule.bas:
  - IsMultiLevelListParagraph
  - ReportAllMultiLevelListLevels
"""
from typing import Protocol, runtime_checkable
from dataclasses import dataclass, field
from tvba_logging import log_event, log_exception


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
class ResolverStatus:
    """Describes the resolver's reliability for callers."""
    mode: str  # "com" | "docx_fallback" | "none"
    reliable_rendered_text: bool
    warnings: list[str] = field(default_factory=list)


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

    def __init__(self, docx_path: str, doc=None):
        self.docx_path = docx_path
        log_event("com_resolver.init.start", path=docx_path, paragraphs=len(doc.paragraphs) if doc is not None else None)
        import win32com.client

        # Use DispatchEx to create a new, independent Word instance.
        # This avoids conflicts with other Word processes and is more
        # reliable in test environments.
        log_event("com_resolver.dispatch.start")
        self.word = win32com.client.DispatchEx("Word.Application")
        log_event("com_resolver.dispatch.done")
        self.word.Visible = False
        try:
            self.word.DisplayAlerts = 0
        except Exception:
            pass
        log_event("com_resolver.documents_open.start", path=self.docx_path)
        self.doc = self.word.Documents.Open(
            self.docx_path,
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
        )
        log_event("com_resolver.documents_open.done")

        # Cache element-to-index only. Caching self.doc.Paragraphs(i) for every
        # paragraph performs one COM round-trip per paragraph and can freeze the
        # GUI on large documents while the progress stays at "Detecting titles".
        self._element_to_index = {}
        if doc is not None:
            for i, para in enumerate(doc.paragraphs):
                self._element_to_index[id(para._element)] = i + 1
        log_event("com_resolver.init.done", indexed=len(self._element_to_index))

    def close(self):
        """Explicitly close Word COM document and quit the application."""
        try:
            if getattr(self, "doc", None):
                log_event("com_resolver.close.doc.start")
                self.doc.Close(SaveChanges=False)
                self.doc = None
                log_event("com_resolver.close.doc.done")
        except Exception:
            log_exception("com_resolver.close.doc.failed")
            pass
        try:
            if getattr(self, "word", None):
                log_event("com_resolver.close.word.start")
                self.word.Quit()
                self.word = None
                log_event("com_resolver.close.word.done")
        except Exception:
            log_exception("com_resolver.close.word.failed")
            pass

    def __enter__(self):
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
        # Fast path: pre-built index mapping.
        cached_idx = self._element_to_index.get(id(para._element))
        if cached_idx is not None:
            return self.doc.Paragraphs(cached_idx)

        # Fallback: compute index dynamically
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


def auto_select(prefer_com: bool = False, docx_path: str | None = None, doc=None) -> tuple[ListResolver, ResolverStatus]:
    """Auto-select best available list resolver.

    Returns (resolver, status) where status describes the resolver's reliability.

    Strategy:
    - If prefer_com=True and Word COM is available: returns ComListResolver (reliable).
    - If prefer_com=True but Word COM fails: returns DocxListResolver with warning.
    - If prefer_com=False: returns DocxListResolver (limited capability noted).
    """
    if prefer_com and docx_path is not None:
        try:
            resolver = ComListResolver(docx_path, doc=doc)
            status = ResolverStatus(
                mode="com",
                reliable_rendered_text=True,
                warnings=[],
            )
            return resolver, status
        except Exception as e:
            log_exception("auto_select.com_failed")
            resolver = DocxListResolver(doc)
            status = ResolverStatus(
                mode="docx_fallback",
                reliable_rendered_text=False,
                warnings=[
                    f"Word COM 不可用，自动编号标题可能无法识别 ({e})",
                ],
            )
            return resolver, status
    resolver = DocxListResolver(doc)
    status = ResolverStatus(
        mode="docx_fallback",
        reliable_rendered_text=False,
        warnings=[],
    )
    return resolver, status

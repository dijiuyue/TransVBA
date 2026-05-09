import tempfile
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest
from docx import Document

from tvba_core_document import apply_settings_to_document
from tvba_settings import FormatSettings


class TestApplySettingsToDocument:
    def test_processes_body_text(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "test.docx"
            doc = Document()
            doc.add_paragraph("正文段落")
            doc.save(path)

            settings = FormatSettings()
            out = apply_settings_to_document(path, settings)
            assert out.exists()

            # Verify output
            result = Document(out)
            assert len(result.paragraphs) == 1

    def test_processes_titles(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "test.docx"
            doc = Document()
            doc.add_paragraph("1 一级标题")
            doc.add_paragraph("1.1 二级标题")
            doc.add_paragraph("正文")
            doc.save(path)

            settings = FormatSettings()
            out = apply_settings_to_document(path, settings)
            result = Document(out)

            # First paragraph should have outline level
            pPr = result.paragraphs[0]._element.find(
                ".//w:pPr",
                {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            )
            outline = pPr.find("w:outlineLvl", {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
            assert outline is not None

    def test_calls_progress_callback(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "test.docx"
            doc = Document()
            doc.add_paragraph("正文")
            doc.save(path)

            progress_calls = []
            def cb(msg, pct):
                progress_calls.append((msg, pct))

            settings = FormatSettings()
            apply_settings_to_document(path, settings, progress_cb=cb)
            assert len(progress_calls) > 0

    def test_custom_output_path(self):
        with tempfile.TemporaryDirectory() as td:
            src = Path(td) / "test.docx"
            out = Path(td) / "output.docx"
            doc = Document()
            doc.add_paragraph("正文")
            doc.save(src)

            settings = FormatSettings()
            result = apply_settings_to_document(src, settings, output_path=out)
            assert result == out
            assert out.exists()

    def test_doc_file_converted_before_processing(self):
        with tempfile.TemporaryDirectory() as td:
            doc_path = Path(td) / "test.doc"
            doc_path.write_text("fake doc content")
            expected_docx = Path(td) / "test.docx"

            # Pre-create the expected .docx so Document() can open it after mocked conversion
            Document().save(expected_docx)

            mock_word = MagicMock()
            mock_doc = MagicMock()
            mock_word.Documents.Open.return_value = mock_doc

            settings = FormatSettings()
            with patch("win32com.client.DispatchEx", return_value=mock_word):
                out = apply_settings_to_document(doc_path, settings)

            assert out.suffix == ".docx"
            assert out.exists()

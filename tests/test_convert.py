"""Tests for tvba_core_convert — .doc to .docx conversion."""
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest
from docx import Document

from tvba_core_convert import ensure_docx


class TestEnsureDocx:
    def test_docx_path_returned_unchanged(self, tmp_path):
        docx = tmp_path / "test.docx"
        Document().save(docx)

        result = ensure_docx(docx)
        assert result == docx

    def test_nonexistent_docx_raises(self, tmp_path):
        missing = tmp_path / "missing.docx"
        with pytest.raises(FileNotFoundError):
            ensure_docx(missing)

    def test_doc_file_triggers_conversion_mocked(self, tmp_path):
        doc = tmp_path / "thesis.doc"
        doc.write_text("fake doc content")
        expected_out = tmp_path / "thesis.docx"

        mock_word = MagicMock()
        mock_doc = MagicMock()
        mock_word.Documents.Open.return_value = mock_doc

        with patch("win32com.client.DispatchEx", return_value=mock_word):
            result = ensure_docx(doc, output_dir=tmp_path)

        mock_word.Documents.Open.assert_called_once_with(str(doc))
        mock_doc.SaveAs2.assert_called_once_with(str(expected_out), FileFormat=16)
        mock_doc.Close.assert_called_once_with(SaveChanges=False)
        mock_word.Quit.assert_called_once()
        assert result == expected_out

    def test_conversion_failure_raises_runtime_error(self, tmp_path):
        doc = tmp_path / "bad.doc"
        doc.write_text("fake")

        with patch(
            "win32com.client.DispatchEx", side_effect=Exception("Word not installed")
        ):
            with pytest.raises(RuntimeError, match="Failed to convert .doc to .docx"):
                ensure_docx(doc, output_dir=tmp_path)

    def test_doc_without_output_dir_uses_same_directory(self, tmp_path):
        doc = tmp_path / "report.doc"
        doc.write_text("fake")
        expected_out = tmp_path / "report.docx"

        mock_word = MagicMock()
        mock_doc = MagicMock()
        mock_word.Documents.Open.return_value = mock_doc

        with patch("win32com.client.DispatchEx", return_value=mock_word):
            result = ensure_docx(doc)

        assert result == expected_out

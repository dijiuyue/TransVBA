import tempfile
from pathlib import Path
import pytest
from docx import Document

from tvba_controller import TvbaController, ValidationResult, ApplyResult
from tvba_persistence import SettingsRepository
from tvba_settings import FormatSettings, BodySettings


class FakeDocumentApplier:
    def __init__(self):
        self.calls = []

    def __call__(self, docx_path, settings, **kwargs):
        self.calls.append((docx_path, settings, kwargs))
        return docx_path


class TestTvbaController:
    def test_settings_property_returns_format_settings(self):
        repo = SettingsRepository()
        applier = FakeDocumentApplier()
        ctrl = TvbaController(repo, applier)
        assert isinstance(ctrl.settings, FormatSettings)

    def test_open_file_sets_opened_file(self):
        repo = SettingsRepository()
        applier = FakeDocumentApplier()
        ctrl = TvbaController(repo, applier)
        ctrl.open_file(Path("C:\\test.docx"))
        assert ctrl.opened_file == Path("C:\\test.docx")

    def test_update_setting_changes_body_font(self):
        repo = SettingsRepository()
        applier = FakeDocumentApplier()
        ctrl = TvbaController(repo, applier)
        result = ctrl.update_setting("body.font", "黑体")
        assert result.valid is True
        assert ctrl.settings.body.font == "黑体"

    def test_update_setting_invalid_path(self):
        repo = SettingsRepository()
        applier = FakeDocumentApplier()
        ctrl = TvbaController(repo, applier)
        result = ctrl.update_setting("invalid.path", "value")
        assert result.valid is False

    def test_apply_calls_applier(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "test.docx"
            doc = Document()
            doc.add_paragraph("test")
            doc.save(path)

            repo = SettingsRepository()
            applier = FakeDocumentApplier()
            ctrl = TvbaController(repo, applier)
            ctrl.open_file(path)
            result = ctrl.apply(save_settings=False)
            assert result.success is True
            assert len(applier.calls) == 1

    def test_reset_to_defaults(self):
        repo = SettingsRepository()
        applier = FakeDocumentApplier()
        ctrl = TvbaController(repo, applier)
        ctrl.update_setting("body.font", "黑体")
        ctrl.reset_to_defaults()
        assert ctrl.settings.body.font == "宋体"

    def test_update_setting_table(self):
        repo = SettingsRepository()
        applier = FakeDocumentApplier()
        ctrl = TvbaController(repo, applier)
        result = ctrl.update_setting("table.title_font", "楷体")
        assert result.valid is True
        assert ctrl.settings.table.title_font == "楷体"

    def test_update_setting_figure(self):
        repo = SettingsRepository()
        applier = FakeDocumentApplier()
        ctrl = TvbaController(repo, applier)
        result = ctrl.update_setting("figure.title_bold", False)
        assert result.valid is True
        assert ctrl.settings.figure.title_bold is False

    def test_update_setting_auto_detect_numeric_titles(self):
        repo = SettingsRepository()
        applier = FakeDocumentApplier()
        ctrl = TvbaController(repo, applier)
        result = ctrl.update_setting("auto_detect_numeric_titles", False)
        assert result.valid is True
        assert ctrl.settings.auto_detect_numeric_titles is False

    def test_update_setting_auto_detect_include_list_paragraphs(self):
        repo = SettingsRepository()
        applier = FakeDocumentApplier()
        ctrl = TvbaController(repo, applier)
        result = ctrl.update_setting("auto_detect_include_list_paragraphs", False)
        assert result.valid is True
        assert ctrl.settings.auto_detect_include_list_paragraphs is False

    def test_update_setting_remember_settings(self):
        repo = SettingsRepository()
        applier = FakeDocumentApplier()
        ctrl = TvbaController(repo, applier)
        result = ctrl.update_setting("remember_settings", False)
        assert result.valid is True
        assert ctrl.settings.remember_settings is False

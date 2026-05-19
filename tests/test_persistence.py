import json
import os
import tempfile
from pathlib import Path
import pytest

from tvba_persistence import SettingsRepository, load_settings, save_settings
from tvba_settings import FormatSettings, BodySettings


class TestSettingsRepository:
    def test_load_missing_returns_defaults(self):
        with tempfile.TemporaryDirectory() as td:
            repo = SettingsRepository(Path(td) / "nonexistent.json")
            settings = repo.load()
            assert isinstance(settings, FormatSettings)
            assert settings.body.font == "宋体"

    def test_save_and_load_roundtrip(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "config.json"
            repo = SettingsRepository(path)
            original = FormatSettings()
            repo.save(original)
            loaded = repo.load()
            assert loaded.body.font == original.body.font
            assert loaded.auto_detect_numeric_titles == original.auto_detect_numeric_titles
            assert len(loaded.titles) == 5

    def test_saves_settings_per_template(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "config.json"
            repo = SettingsRepository(path)
            repo.save(FormatSettings(template_name="general_spec", body=BodySettings(font="黑体")))
            repo.save(FormatSettings(template_name="dapeng_internal", body=BodySettings(font="楷体")))

            assert repo.load_for_template("general_spec").body.font == "黑体"
            assert repo.load_for_template("dapeng_internal").body.font == "楷体"

    def test_clear_one_template_keeps_others(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "config.json"
            repo = SettingsRepository(path)
            repo.save(FormatSettings(template_name="general_spec", body=BodySettings(font="黑体")))
            repo.save(FormatSettings(template_name="dapeng_internal", body=BodySettings(font="楷体")))

            repo.clear("dapeng_internal")

            assert repo.load_for_template("dapeng_internal") == FormatSettings()
            assert repo.load_for_template("general_spec").body.font == "黑体"

    def test_legacy_single_template_file_loads_matching_template(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "config.json"
            path.write_text(
                json.dumps({"template_name": "dapeng_internal", "body": {"font": "楷体"}}),
                encoding="utf-8",
            )
            repo = SettingsRepository(path)

            assert repo.load_for_template("dapeng_internal").body.font == "楷体"
            assert repo.load_for_template("general_spec") == FormatSettings()

    def test_load_corrupt_returns_defaults(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "config.json"
            path.write_text("not json{{{", encoding="utf-8")
            repo = SettingsRepository(path)
            settings = repo.load()
            assert isinstance(settings, FormatSettings)
            assert settings.body.font == "宋体"

    def test_save_creates_directories(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "deep" / "nested" / "config.json"
            repo = SettingsRepository(path)
            repo.save(FormatSettings())
            assert path.exists()


class TestModuleFunctions:
    def test_load_settings_default_path(self, monkeypatch):
        with tempfile.TemporaryDirectory() as td:
            fake_appdata = Path(td)
            monkeypatch.setenv("APPDATA", str(fake_appdata))
            settings = load_settings()
            assert isinstance(settings, FormatSettings)

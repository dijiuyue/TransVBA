"""JSON persistence for format settings.

Corresponds to VBA FormatModule.bas:
  - LoadSettingsFromRegistry
  - SaveSettingsToRegistry
"""
import json
import os
from dataclasses import asdict
from pathlib import Path

from tvba_settings import FormatSettings

DEFAULT_CONFIG_DIR = Path(os.environ.get("APPDATA", "")) / "TransVBA"
DEFAULT_CONFIG_PATH = DEFAULT_CONFIG_DIR / "tvba_config.json"


class SettingsRepository:
    def __init__(self, path: Path = DEFAULT_CONFIG_PATH):
        self.path = path

    def load(self) -> FormatSettings:
        if not self.path.exists():
            return FormatSettings()
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return self._from_dict(data)
        except (json.JSONDecodeError, OSError, TypeError, KeyError):
            return FormatSettings()

    def save(self, settings: FormatSettings) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(asdict(settings), f, ensure_ascii=False, indent=2)

    def _from_dict(self, data: dict) -> FormatSettings:
        from tvba_settings import (
            BodySettings, TitleLevelSettings, TableSettings,
            FigureSettings, TocLegacyFixedDefaults, ValidationRules,
            CoverSettings, AppendixSettings, HeaderSettings,
        )
        titles = tuple(
            TitleLevelSettings(**t) for t in data.get("titles", [])
        )
        if len(titles) != 5:
            titles = tuple(TitleLevelSettings() for _ in range(5))
        body_data = data.get("body", {})
        # Backward compat: old configs used special_indent_cm (cm) instead of chars
        if "special_indent_cm" in body_data and "special_indent_chars" not in body_data:
            from tvba_utils import cm_to_points
            body_data["special_indent_chars"] = cm_to_points(body_data.pop("special_indent_cm")) / 12.0

        validation_data = data.get("validation", {})
        forbidden = validation_data.pop("forbidden_words", [])
        if isinstance(forbidden, list):
            validation_data["forbidden_words"] = tuple(forbidden)

        return FormatSettings(
            template_name=data.get("template_name", "general_spec"),
            validation=ValidationRules(**validation_data),
            body=BodySettings(**body_data),
            titles=titles,
            table=_table_from_dict(data.get("table", {})),
            figure=FigureSettings(**data.get("figure", {})),
            toc=TocLegacyFixedDefaults(**data.get("toc", {})),
            cover=CoverSettings(**data.get("cover", {})),
            appendix=AppendixSettings(**data.get("appendix", {})),
            header=HeaderSettings(**data.get("header", {})),
            auto_detect_numeric_titles=data.get("auto_detect_numeric_titles", True),
            auto_detect_include_list_paragraphs=data.get("auto_detect_include_list_paragraphs", True),
            remember_settings=data.get("remember_settings", True),
            prefer_com_resolver=data.get("prefer_com_resolver", False),
        )


def _table_from_dict(data: dict) -> "TableSettings":
    """Build TableSettings from dict with backward compat for auto_fit_window."""
    from tvba_settings import TableSettings
    if "auto_fit_mode" not in data:
        old_val = data.pop("auto_fit_window", True)
        data["auto_fit_mode"] = "window" if old_val else "fixed"
    return TableSettings(**data)


def load_settings(path: Path | None = None) -> FormatSettings:
    repo = SettingsRepository(path or DEFAULT_CONFIG_PATH)
    return repo.load()


def save_settings(settings: FormatSettings, path: Path | None = None) -> None:
    repo = SettingsRepository(path or DEFAULT_CONFIG_PATH)
    repo.save(settings)

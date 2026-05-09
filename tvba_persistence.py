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
            FigureSettings, TocLegacyFixedDefaults,
        )
        titles = tuple(
            TitleLevelSettings(**t) for t in data.get("titles", [])
        )
        if len(titles) != 5:
            titles = tuple(TitleLevelSettings() for _ in range(5))
        return FormatSettings(
            body=BodySettings(**data.get("body", {})),
            titles=titles,
            table=TableSettings(**data.get("table", {})),
            figure=FigureSettings(**data.get("figure", {})),
            toc=TocLegacyFixedDefaults(**data.get("toc", {})),
            auto_detect_numeric_titles=data.get("auto_detect_numeric_titles", True),
            auto_detect_include_list_paragraphs=data.get("auto_detect_include_list_paragraphs", True),
            remember_settings=data.get("remember_settings", True),
            prefer_com_resolver=data.get("prefer_com_resolver", False),
        )


def load_settings(path: Path | None = None) -> FormatSettings:
    repo = SettingsRepository(path or DEFAULT_CONFIG_PATH)
    return repo.load()


def save_settings(settings: FormatSettings, path: Path | None = None) -> None:
    repo = SettingsRepository(path or DEFAULT_CONFIG_PATH)
    repo.save(settings)

"""Controller layer — mediates between View (Tkinter) and Model (settings + core).

Completely independent of Tkinter. Testable with mock applier.
"""
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Any, Callable, Protocol

from tvba_persistence import SettingsRepository
from tvba_settings import FormatSettings, BodySettings, TitleLevelSettings


@dataclass
class ValidationResult:
    valid: bool
    message: str = ""


@dataclass
class ApplyResult:
    success: bool
    message: str = ""
    output_path: Path | None = None


class DocumentApplier(Protocol):
    def __call__(
        self,
        docx_path: Path,
        settings: FormatSettings,
        *,
        output_path: Path | None = None,
        progress_cb: Callable | None = None,
    ) -> Path:
        ...


class TvbaController:
    def __init__(self, repo: SettingsRepository, applier: DocumentApplier):
        self._repo = repo
        self._applier = applier
        self._settings = repo.load()
        self._opened_file: Path | None = None

    @property
    def settings(self) -> FormatSettings:
        return self._settings

    @property
    def opened_file(self) -> Path | None:
        return self._opened_file

    def open_file(self, path: Path) -> None:
        self._opened_file = path

    def update_setting(self, path: str, value: Any) -> ValidationResult:
        """Update a setting by dotted path like 'body.font' or 'titles.0.size'."""
        parts = path.split(".")
        try:
            if parts[0] == "body" and len(parts) == 2:
                attr = parts[1]
                new_body = replace(self._settings.body, **{attr: value})
                self._settings = replace(
                    self._settings,
                    body=new_body,
                )
                return ValidationResult(valid=True)

            elif parts[0] == "titles" and len(parts) == 3:
                idx = int(parts[1])
                attr = parts[2]
                current = self._settings.titles[idx]
                new_title = replace(current, **{attr: value})
                new_titles = list(self._settings.titles)
                new_titles[idx] = new_title
                self._settings = replace(
                    self._settings,
                    titles=tuple(new_titles),
                )
                return ValidationResult(valid=True)

            else:
                return ValidationResult(valid=False, message=f"Unknown path: {path}")
        except (AttributeError, IndexError, TypeError) as e:
            return ValidationResult(valid=False, message=str(e))

    def apply(self, *, save_settings: bool, progress_cb=None) -> ApplyResult:
        if self._opened_file is None:
            return ApplyResult(success=False, message="No file opened")

        try:
            out = self._applier(
                self._opened_file,
                self._settings,
                progress_cb=progress_cb,
            )
            if save_settings and self._settings.remember_settings:
                self._repo.save(self._settings)
            return ApplyResult(success=True, output_path=out)
        except Exception as e:
            return ApplyResult(success=False, message=str(e))

    def reset_to_defaults(self) -> None:
        self._settings = FormatSettings()

    def load_preset(self, name: str) -> None:
        # TODO: Implement preset loading from JSON files
        pass

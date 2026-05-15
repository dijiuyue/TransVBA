"""Controller layer — mediates between View (Tkinter) and Model (settings + core).

Completely independent of Tkinter. Testable with mock applier.
"""
import time
import traceback
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
    elapsed_ms: int = 0


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
    def __init__(self, repo: SettingsRepository, applier: DocumentApplier,
                 default_settings: FormatSettings | None = None):
        self._repo = repo
        self._applier = applier
        template = default_settings or FormatSettings()
        self._template_defaults = template
        self._settings = template
        self._opened_file: Path | None = None

    @property
    def settings(self) -> FormatSettings:
        return self._settings

    @property
    def opened_file(self) -> Path | None:
        return self._opened_file

    def open_file(self, path: Path) -> None:
        self._opened_file = path

    def switch_template(self, template_id: str) -> FormatSettings:
        from tvba_templates import TemplateManager
        self._settings = TemplateManager.load_template(template_id)
        return self._settings

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

            elif parts[0] == "table" and len(parts) == 2:
                attr = parts[1]
                new_table = replace(self._settings.table, **{attr: value})
                self._settings = replace(
                    self._settings,
                    table=new_table,
                )
                return ValidationResult(valid=True)

            elif parts[0] == "figure" and len(parts) == 2:
                attr = parts[1]
                new_figure = replace(self._settings.figure, **{attr: value})
                self._settings = replace(
                    self._settings,
                    figure=new_figure,
                )
                return ValidationResult(valid=True)

            elif parts[0] in ("auto_detect_numeric_titles", "auto_detect_include_list_paragraphs", "remember_settings", "prefer_com_resolver", "template_name") and len(parts) == 1:
                self._settings = replace(
                    self._settings,
                    **{parts[0]: value},
                )
                return ValidationResult(valid=True)

            else:
                return ValidationResult(valid=False, message=f"Unknown path: {path}")
        except (AttributeError, IndexError, TypeError) as e:
            return ValidationResult(valid=False, message=str(e))

    def apply(self, *, save_settings: bool, progress_cb=None) -> ApplyResult:
        if self._opened_file is None:
            return ApplyResult(success=False, message="No file opened")

        stem = self._opened_file.stem
        suffix = self._opened_file.suffix
        output_path = self._opened_file.parent / f"{stem}+格式修改后{suffix}"

        start = time.perf_counter()
        try:
            out = self._applier(
                self._opened_file,
                self._settings,
                output_path=output_path,
                progress_cb=progress_cb,
            )
            elapsed_ms = int((time.perf_counter() - start) * 1000)
            if save_settings and self._settings.remember_settings:
                self._repo.save(self._settings)
            return ApplyResult(success=True, output_path=out, elapsed_ms=elapsed_ms)
        except Exception as e:
            elapsed_ms = int((time.perf_counter() - start) * 1000)
            tb = traceback.format_exc()
            return ApplyResult(success=False, message=f"{e}\n\n{tb}", elapsed_ms=elapsed_ms)

    def reset_to_template_defaults(self) -> None:
        self._settings = self._template_defaults

    def clear_saved_settings(self) -> None:
        self._repo.clear()

    def load_saved_settings(self) -> None:
        saved = self._repo.load()
        if saved != FormatSettings():
            self._settings = saved

    def load_preset(self, name: str) -> None:
        # TODO: Implement preset loading from JSON files
        pass

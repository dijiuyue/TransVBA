"""Controller layer — mediates between View (Tkinter) and Model (settings + core).

Completely independent of Tkinter. Testable with mock applier.
"""
import json
import time
import traceback
from dataclasses import asdict, dataclass, field, replace
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Protocol

from tvba_persistence import SettingsRepository
from tvba_settings import FormatSettings, BodySettings, TitleLevelSettings

PRESETS_DIR = Path(__file__).parent / "presets"


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
    warnings: list[str] = field(default_factory=list)


class DocumentApplier(Protocol):
    def __call__(
        self,
        docx_path: Path,
        settings: FormatSettings,
        *,
        output_path: Path | None = None,
        progress_cb: Callable | None = None,
    ) -> tuple[Path, object] | Path:
        ...



class TvbaController:
    def __init__(self, repo: SettingsRepository, applier: DocumentApplier,
                 default_settings: FormatSettings | None = None):
        self._repo = repo
        self._applier = applier
        template = default_settings or FormatSettings()
        self._template_defaults = template
        self._settings = template
        self._current_template_id = template.template_name
        self._opened_file: Path | None = None

    @property
    def settings(self) -> FormatSettings:
        return self._settings

    @property
    def opened_file(self) -> Path | None:
        return self._opened_file

    @property
    def current_template_id(self) -> str:
        return self._current_template_id

    def is_custom_template(self) -> bool:
        return self._current_template_id == "__custom__"

    def open_file(self, path: Path) -> None:
        self._opened_file = path

    def switch_template(self, template_id: str) -> FormatSettings:
        from tvba_templates import TemplateManager
        self._current_template_id = template_id
        base = TemplateManager.load_template(template_id)
        self._template_defaults = base
        self._settings = base
        # Load saved overrides for this template
        saved = self._repo.load_for_template(template_id)
        if saved != FormatSettings() and saved.template_name == template_id:
            self._settings = saved
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

        output_path = self._make_output_path()

        # Check if output file is in use (locked by Word or another process)
        if output_path.exists():
            if not _can_write_file(output_path):
                return ApplyResult(
                    success=False,
                    message=f"无法写入输出文件:\n{output_path}\n\n"
                            f"文件可能正在被 Word 或其他程序占用。\n请关闭该文件后重试。"
                )

        start = time.perf_counter()
        try:
            result = self._applier(
                self._opened_file,
                self._settings,
                output_path=output_path,
                progress_cb=progress_cb,
            )
            elapsed_ms = int((time.perf_counter() - start) * 1000)
            if save_settings and self._settings.remember_settings:
                settings_to_save = replace(self._settings, template_name=self._current_template_id)
                self._repo.save(settings_to_save)
            # Handle new (path, warnings) tuple or legacy path-only return
            if isinstance(result, tuple):
                out, warnings_obj = result
                warning_msgs = warnings_obj.messages if hasattr(warnings_obj, 'messages') else []
                return ApplyResult(success=True, output_path=out, elapsed_ms=elapsed_ms, warnings=warning_msgs)
            else:
                return ApplyResult(success=True, output_path=result, elapsed_ms=elapsed_ms)
        except PermissionError:
            elapsed_ms = int((time.perf_counter() - start) * 1000)
            return ApplyResult(
                success=False,
                message=f"无法写入输出文件:\n{output_path}\n\n"
                        f"文件可能正在被 Word 或其他程序占用。\n请关闭该文件后重试。",
                elapsed_ms=elapsed_ms,
            )
        except Exception as e:
            elapsed_ms = int((time.perf_counter() - start) * 1000)
            tb = traceback.format_exc()
            return ApplyResult(success=False, message=f"{e}\n\n{tb}", elapsed_ms=elapsed_ms)

    def reset_to_template_defaults(self) -> None:
        self._settings = self._template_defaults

    def clear_saved_settings(self) -> None:
        self._repo.clear(self._current_template_id)

    def load_saved_settings(self) -> None:
        saved = self._repo.load_for_template(self._current_template_id)
        if saved != FormatSettings() and saved.template_name == self._current_template_id:
            self._settings = self._migrate_saved_settings(saved)

    def _migrate_saved_settings(self, saved: FormatSettings) -> FormatSettings:
        if saved.template_name != "dapeng_internal" or len(saved.titles) != 5:
            return saved

        saved_bolds = tuple(title.bold for title in saved.titles)
        legacy_bold_patterns = {
            (True, True, True, False, False),
            (True, True, True, True, True),
        }
        has_legacy_title_indent = any(
            title.left_indent_chars != self._template_defaults.titles[i].left_indent_chars
            or title.right_indent_chars != self._template_defaults.titles[i].right_indent_chars
            or title.special_indent != self._template_defaults.titles[i].special_indent
            or title.special_indent_chars != self._template_defaults.titles[i].special_indent_chars
            for i, title in enumerate(saved.titles)
        )
        if saved_bolds not in legacy_bold_patterns and not has_legacy_title_indent:
            return saved

        titles = tuple(
            replace(
                title,
                bold=self._template_defaults.titles[i].bold if saved_bolds in legacy_bold_patterns else title.bold,
                left_indent_chars=self._template_defaults.titles[i].left_indent_chars,
                right_indent_chars=self._template_defaults.titles[i].right_indent_chars,
                special_indent=self._template_defaults.titles[i].special_indent,
                special_indent_chars=self._template_defaults.titles[i].special_indent_chars,
            )
            for i, title in enumerate(saved.titles)
        )
        return replace(saved, titles=titles)

    def load_preset(self, name: str) -> bool:
        """Load settings from a named preset JSON file. Returns True on success."""
        if not PRESETS_DIR.exists():
            return False
        preset_path = PRESETS_DIR / f"{name}.json"
        if not preset_path.exists():
            return False
        try:
            with open(preset_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            from tvba_templates import _settings_from_dict
            self._settings = _settings_from_dict(data)
            return True
        except (json.JSONDecodeError, OSError, KeyError):
            return False

    def save_preset(self, name: str) -> bool:
        """Save current settings as a named preset JSON file. Returns True on success."""
        PRESETS_DIR.mkdir(parents=True, exist_ok=True)
        preset_path = PRESETS_DIR / f"{name}.json"
        try:
            data = asdict(self._settings)
            with open(preset_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            return True
        except OSError:
            return False

    @staticmethod
    def list_presets() -> list[str]:
        """List available preset names."""
        if not PRESETS_DIR.exists():
            return []
        return sorted(f.stem for f in PRESETS_DIR.glob("*.json"))

    def get_all_template_ids(self) -> list[str]:
        """Return all template IDs (file-based + custom) for multi-template validation."""
        from tvba_templates import TemplateManager
        return [t["id"] for t in TemplateManager.list_templates()]

    def load_template_for_validation(self, template_id: str) -> FormatSettings:
        """Load a template for validation, merging saved overrides if any."""
        from tvba_templates import TemplateManager
        base = TemplateManager.load_template(template_id)
        saved = self._repo.load_for_template(template_id)
        if saved != FormatSettings() and saved.template_name == template_id:
            return saved
        return base

    def _make_output_path(self) -> Path:
        """Generate a safe output path with timestamp collision avoidance.

        Always uses .docx suffix regardless of input format, since ensure_docx()
        converts .doc to OOXML format before processing.
        """
        stem = self._opened_file.stem
        suffix = ".docx"
        parent = self._opened_file.parent
        candidate = parent / f"{stem}+格式修改后{suffix}"
        if candidate.exists():
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            candidate = parent / f"{stem}+格式修改后_{ts}{suffix}"
        return candidate


def _can_write_file(path: Path) -> bool:
    """Check if a file can be written to (not locked by another process)."""
    try:
        with open(path, "a"):
            pass
        return True
    except (PermissionError, OSError):
        return False

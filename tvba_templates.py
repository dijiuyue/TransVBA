"""Template manager — load and manage template configurations from JSON files.

Each template is a standalone JSON file under templates/ containing a full
FormatSettings structure plus template metadata and validation rules.
"""
import json
from pathlib import Path
from tvba_settings import FormatSettings


TEMPLATES_DIR = Path(__file__).parent / "templates"


class TemplateManager:
    """Load and manage template configurations from JSON files."""

    @classmethod
    def list_templates(cls) -> list[dict]:
        """Scan templates/ dir, return list of {id, name, description}."""
        if not TEMPLATES_DIR.exists():
            return []
        templates = []
        for f in sorted(TEMPLATES_DIR.glob("*.json")):
            try:
                with open(f, "r", encoding="utf-8") as fh:
                    data = json.load(fh)
                templates.append({
                    "id": data.get("template_name", f.stem),
                    "name": data.get("template_display", f.stem),
                    "description": data.get("template_description", ""),
                })
            except (json.JSONDecodeError, OSError):
                continue
        return templates

    @classmethod
    def load_template(cls, template_id: str) -> FormatSettings:
        """Load a specific template and return FormatSettings."""
        file_path = TEMPLATES_DIR / f"{template_id}.json"
        if not file_path.exists():
            raise FileNotFoundError(f"Template not found: {template_id}")
        with open(file_path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return _settings_from_dict(data)

    @classmethod
    def get_default_template_id(cls) -> str:
        """Return the default template ID."""
        templates = cls.list_templates()
        if templates:
            return templates[0]["id"]
        return "general_spec"


def _settings_from_dict(data: dict) -> FormatSettings:
    """Build FormatSettings from a dict (shared with persistence layer)."""
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
    if "special_indent_cm" in body_data and "special_indent_chars" not in body_data:
        from tvba_utils import cm_to_points
        body_data["special_indent_chars"] = cm_to_points(body_data.pop("special_indent_cm")) / 12.0

    validation_data = data.get("validation", {})
    forbidden = validation_data.pop("forbidden_words", [])
    if isinstance(forbidden, list):
        validation_data["forbidden_words"] = tuple(forbidden)

    table_data = data.get("table", {})
    if "auto_fit_mode" not in table_data:
        old_val = table_data.pop("auto_fit_window", True) if isinstance(table_data, dict) else True
        table_data["auto_fit_mode"] = "window" if old_val else "fixed"

    return FormatSettings(
        template_name=data.get("template_name", "general_spec"),
        validation=ValidationRules(**validation_data),
        body=BodySettings(**body_data),
        titles=titles,
        table=TableSettings(**table_data),
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

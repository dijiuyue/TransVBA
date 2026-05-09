"""Data classes for format settings.

Corresponds to VBA Typemodule.bas: FormatSettings and related types.
"""
from dataclasses import dataclass, field


@dataclass(frozen=True)
class TitleLevelSettings:
    """Single title level settings."""
    alignment: str = "左对齐"
    font: str = "宋体"
    size: str = "小四"
    bold: bool = False
    before_lines: float = 0.5
    after_lines: float = 0.5
    line_spacing: float = 1.5


@dataclass(frozen=True)
class BodySettings:
    """Body text formatting settings."""
    font: str = "宋体"
    size: str = "小四"
    spacing: float = 1.5
    before_lines: float = 0.0
    after_lines: float = 0.0
    alignment: str = "两端对齐"
    left_indent_cm: float = 0.0
    right_indent_cm: float = 0.0
    special_indent: str = "首行缩进"
    special_indent_chars: float = 2.0


@dataclass(frozen=True)
class TableSettings:
    """Table and table caption formatting settings."""
    title_font: str = "黑体"
    title_size: str = "小四"
    title_bold: bool = True
    title_spacing: float = 1.5
    title_left_indent_cm: float = 0.0
    title_right_indent_cm: float = 0.0
    title_special_indent: str = "无"
    title_special_indent_cm: float = 0.0
    body_font: str = "宋体"
    body_size: str = "五号"
    line_width_pt: float = 0.5
    row_height_cm: float = 0.7
    spacing: float = 1.0
    auto_fit_window: bool = True


@dataclass(frozen=True)
class FigureSettings:
    """Figure caption formatting settings."""
    title_font: str = "黑体"
    title_size: str = "小四"
    title_bold: bool = True
    title_spacing: float = 1.5
    title_left_indent_cm: float = 0.0
    title_right_indent_cm: float = 0.0
    title_special_indent: str = "无"
    title_special_indent_cm: float = 0.0


@dataclass(frozen=True)
class TocLegacyFixedDefaults:
    """Fixed TOC defaults (not exposed in UI, but kept for 1:1 VBA equivalence)."""
    title_font: str = "宋体"
    title_size: str = "小四"
    title_bold: bool = True
    title_spacing: float = 1.5
    level1_font: str = "宋体"
    level1_size: str = "小四"
    level1_bold: bool = True
    level2_font: str = "宋体"
    level2_size: str = "小四"
    level2_indent_chars: int = 2
    level3_font: str = "宋体"
    level3_size: str = "小四"
    level3_indent_chars: int = 4


@dataclass
class FormatSettings:
    """Complete format settings — corresponds to VBA FormatSettings."""
    body: BodySettings = field(default_factory=BodySettings)
    titles: tuple[TitleLevelSettings, ...] = field(
        default_factory=lambda: tuple(TitleLevelSettings() for _ in range(5))
    )
    table: TableSettings = field(default_factory=TableSettings)
    figure: FigureSettings = field(default_factory=FigureSettings)
    toc: TocLegacyFixedDefaults = field(default_factory=TocLegacyFixedDefaults)

    auto_detect_numeric_titles: bool = True
    auto_detect_include_list_paragraphs: bool = True
    remember_settings: bool = True
    prefer_com_resolver: bool = False

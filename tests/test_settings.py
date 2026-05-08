import pytest
from dataclasses import asdict
from tvba_settings import (
    TitleLevelSettings,
    BodySettings,
    TableSettings,
    FigureSettings,
    TocLegacyFixedDefaults,
    FormatSettings,
)


class TestTitleLevelSettings:
    def test_defaults(self):
        t = TitleLevelSettings()
        assert t.alignment == "左对齐"
        assert t.font == "宋体"
        assert t.size == "小四"
        assert t.bold is False
        assert t.before_lines == 0.5
        assert t.after_lines == 0.5
        assert t.line_spacing == 1.5

    def test_frozen_cannot_mutate(self):
        t = TitleLevelSettings()
        with pytest.raises(AttributeError):
            t.font = "黑体"


class TestBodySettings:
    def test_defaults(self):
        b = BodySettings()
        assert b.font == "宋体"
        assert b.size == "小四"
        assert b.spacing == 1.5
        assert b.special_indent == "首行缩进"
        assert b.special_indent_cm == 0.74

    def test_frozen_cannot_mutate(self):
        b = BodySettings()
        with pytest.raises(AttributeError):
            b.font = "黑体"


class TestFormatSettings:
    def test_defaults(self):
        fs = FormatSettings()
        assert isinstance(fs.body, BodySettings)
        assert len(fs.titles) == 5
        assert all(isinstance(t, TitleLevelSettings) for t in fs.titles)
        assert fs.auto_detect_numeric_titles is True
        assert fs.auto_detect_include_list_paragraphs is True
        assert fs.remember_settings is True

    def test_titles_are_independent(self):
        fs = FormatSettings()
        assert fs.titles[0] is not fs.titles[1]

    def test_asdict_roundtrip_structure(self):
        fs = FormatSettings()
        d = asdict(fs)
        assert "body" in d
        assert "titles" in d
        assert len(d["titles"]) == 5
        assert d["auto_detect_numeric_titles"] is True

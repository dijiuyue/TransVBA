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
        assert t.left_indent_chars == 0.0
        assert t.right_indent_chars == 0.0
        assert t.special_indent == "无"
        assert t.special_indent_chars == 0.0
        assert t.normalize_brackets is False

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
        assert b.special_indent_chars == 2.0

    def test_frozen_cannot_mutate(self):
        b = BodySettings()
        with pytest.raises(AttributeError):
            b.font = "黑体"


class TestCoverSettings:
    def test_defaults(self):
        from tvba_settings import CoverSettings
        c = CoverSettings()
        assert c.font == "宋体"
        assert c.size == "二号"
        assert c.bold is True
        assert c.line_spacing == 1.5
        assert c.alignment == "居中"

    def test_frozen_cannot_mutate(self):
        from tvba_settings import CoverSettings
        c = CoverSettings()
        with pytest.raises(AttributeError):
            c.font = "黑体"


class TestAppendixSettings:
    def test_defaults(self):
        from tvba_settings import AppendixSettings
        a = AppendixSettings()
        assert a.title_font == "宋体"
        assert a.title_size == "小四"
        assert a.title_bold is True
        assert a.title_line_spacing == 1.5
        assert a.body_font == "宋体"
        assert a.body_size == "小五"
        assert a.body_bold is False
        assert a.body_line_spacing == 1.0

    def test_frozen_cannot_mutate(self):
        from tvba_settings import AppendixSettings
        a = AppendixSettings()
        with pytest.raises(AttributeError):
            a.title_font = "黑体"


class TestFigureSettings:
    def test_defaults(self):
        f = FigureSettings()
        assert f.title_bold is False


class TestHeaderSettings:
    def test_defaults(self):
        from tvba_settings import HeaderSettings
        h = HeaderSettings()
        assert h.font == "宋体"
        assert h.size == "小五"
        assert h.bold is False
        assert h.line_spacing == 1.0

    def test_frozen_cannot_mutate(self):
        from tvba_settings import HeaderSettings
        h = HeaderSettings()
        with pytest.raises(AttributeError):
            h.font = "黑体"


class TestFormatSettings:
    def test_defaults(self):
        fs = FormatSettings()
        assert isinstance(fs.body, BodySettings)
        assert len(fs.titles) == 5
        assert all(isinstance(t, TitleLevelSettings) for t in fs.titles)
        assert fs.auto_detect_numeric_titles is True
        assert fs.auto_detect_include_list_paragraphs is True
        assert fs.remember_settings is True

    def test_has_cover_appendix_header(self):
        from tvba_settings import CoverSettings, AppendixSettings, HeaderSettings
        fs = FormatSettings()
        assert isinstance(fs.cover, CoverSettings)
        assert isinstance(fs.appendix, AppendixSettings)
        assert isinstance(fs.header, HeaderSettings)

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
        assert "cover" in d
        assert "appendix" in d
        assert "header" in d

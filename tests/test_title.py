import pytest
from docx import Document

from tvba_core_title import (
    identify_numeric_title_level,
    identify_level_from_number,
    normalize_number_string,
    apply_title_style,
)
from tvba_settings import TitleLevelSettings, BodySettings


class TestNormalizeNumberString:
    def test_fullwidth_dot_to_halfwidth(self):
        assert normalize_number_string("1．2．3") == "1.2.3"

    def test_removes_trailing_dot(self):
        assert normalize_number_string("1.2.") == "1.2"

    def test_strips_whitespace(self):
        assert normalize_number_string("  1.1  ") == "1.1"

    def test_multiple_fullwidth_dots(self):
        assert normalize_number_string("１．２．３．４") == "1.2.3.4"


class TestIdentifyLevelFromNumber:
    def test_level_1_no_dot(self):
        assert identify_level_from_number("1") == 1

    def test_level_1_with_dot_zero(self):
        assert identify_level_from_number("1.0") == 1

    def test_level_2_one_dot(self):
        assert identify_level_from_number("1.1") == 2

    def test_level_3_two_dots(self):
        assert identify_level_from_number("1.1.2") == 3

    def test_level_4_three_dots(self):
        assert identify_level_from_number("1.1.2.3") == 4

    def test_level_5_four_dots(self):
        assert identify_level_from_number("1.1.2.3.4") == 5

    def test_level_0_too_many_dots(self):
        assert identify_level_from_number("1.1.2.3.4.5") == 0

    def test_level_0_empty(self):
        assert identify_level_from_number("") == 0


class TestIdentifyNumericTitleLevel:
    def test_level_1_simple(self):
        assert identify_numeric_title_level("1 引言") == 1

    def test_level_2_simple(self):
        assert identify_numeric_title_level("1.1 背景") == 2

    def test_level_3_simple(self):
        assert identify_numeric_title_level("1.1.1 详细背景") == 3

    def test_no_number_returns_0(self):
        assert identify_numeric_title_level("引言") == 0

    def test_number_without_space_returns_0(self):
        assert identify_numeric_title_level("1引言") == 0

    def test_fullwidth_dots_work(self):
        assert identify_numeric_title_level("1．1 背景") == 2

    def test_tab_separator_works(self):
        assert identify_numeric_title_level("1.1\t背景") == 2

    def test_trailing_dot_normalized(self):
        assert identify_numeric_title_level("1. 引言") == 1

    def test_level_1_requires_space_or_tab(self):
        assert identify_numeric_title_level("1引言") == 0
        assert identify_numeric_title_level("1 引言") == 1

    def test_too_many_dots_returns_0(self):
        assert identify_numeric_title_level("1.1.2.3.4.5 太深") == 0

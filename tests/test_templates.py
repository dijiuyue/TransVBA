from tvba_templates import TemplateManager


class TestTemplateSpecificDefaults:
    def test_general_level4_and_table_defaults(self):
        settings = TemplateManager.load_template("general_spec")
        level4 = settings.titles[3]

        assert level4.alignment == "左对齐"
        assert level4.left_indent_chars == 2.0
        assert level4.special_indent == "无"
        assert level4.normalize_brackets is True
        assert settings.table.row_height_cm == 0.6
        assert settings.table.auto_fit_mode == "window"
        assert settings.table.line_width_pt == 0.25

    def test_dapeng_defaults_unchanged(self):
        settings = TemplateManager.load_template("dapeng_internal")
        level4 = settings.titles[3]

        assert level4.left_indent_chars == 0.0
        assert level4.normalize_brackets is False
        assert settings.table.row_height_cm == 0.0
        assert settings.table.line_width_pt == 0.5

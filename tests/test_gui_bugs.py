"""Tests for GUI bugs that cause "Apply" button to appear to do nothing.

Root cause chain:
1. Panels are lazy-built — _populate_from_settings runs in __init__ before panels exist
2. When user clicks a tree node, panel is built but values are empty
3. _sync_settings_to_controller calls float('') → ValueError
4. _sync_settings_to_controller is OUTSIDE the try/except in _on_apply
5. Tkinter silently swallows the exception → no visible feedback
"""
import tkinter as tk
from unittest.mock import MagicMock, patch

import pytest

from tvba_gui import TvbaMainWindow
from tvba_settings import FormatSettings


class TestGuiLazyPanelBug:
    @pytest.fixture(scope="class", autouse=True)
    def _setup_window(self, request):
        ctrl = MagicMock()
        ctrl.settings = FormatSettings()
        root = TvbaMainWindow(ctrl)
        root.withdraw()
        request.cls.window = root
        yield
        try:
            root.destroy()
        except tk.TclError:
            pass

    def test_body_panel_populated_after_first_select(self):
        """Panel values must be filled from controller settings when first built."""
        assert "body" not in self.window._panels

        self.window.tree.selection_set("body")
        self.window._on_tree_select(None)

        assert "body" in self.window._panels
        assert self.window.cmb_body_font.get() == "宋体"
        assert self.window.spn_body_spacing.get() == "1.5"
        assert self.window.cmb_body_align.get() == "两端对齐"

    def test_title_panel_populated_after_first_select(self):
        self.window.tree.selection_set("title_1")
        self.window._on_tree_select(None)

        assert "title_1" in self.window._panels
        assert self.window.cmb_title_1_font.get() == "宋体"
        assert self.window.spn_title_1_spacing.get() == "1.5"

    def test_sync_settings_does_not_crash_on_empty_spinbox(self):
        """Empty spinbox values should not crash _sync_settings_to_controller."""
        self.window.tree.selection_set("body")
        self.window._on_tree_select(None)

        # Simulate user clearing a spinbox
        self.window.spn_body_spacing.set("")

        # Should not raise
        self.window._sync_settings_to_controller()

    def test_on_apply_catches_sync_exceptions(self):
        """_on_apply must catch exceptions from _sync_settings_to_controller."""
        self.window.tree.selection_set("body")
        self.window._on_tree_select(None)

        # Sync is gated by 修改模式; must enable it for sync to run
        self.window.chk_edit.set(True)

        # Force _sync_settings_to_controller to raise even though it is now defensive
        with patch.object(self.window, "_sync_settings_to_controller", side_effect=RuntimeError("boom")):
            with patch("tvba_gui.messagebox") as mock_mb:
                # Should not propagate an uncaught exception
                self.window._on_apply()

                # Error dialog should be shown
                mock_mb.showerror.assert_called_once()
                args = mock_mb.showerror.call_args[0]
                assert "同步" in args[1] or "boom" in args[1]

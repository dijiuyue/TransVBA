"""TransVBA entry point.

Corresponds to VBA Mainmodule.bas: ShowFormatSettings
"""
import tkinter as tk
from pathlib import Path

from tvba_persistence import SettingsRepository
from tvba_controller import TvbaController
from tvba_core_document import apply_settings_to_document
from tvba_gui import TvbaMainWindow
from tvba_settings import FormatSettings
from tvba_templates import TemplateManager


def main():
    repo = SettingsRepository()
    try:
        default_settings = TemplateManager.load_template(
            TemplateManager.get_default_template_id()
        )
    except (FileNotFoundError, Exception):
        default_settings = FormatSettings()
    controller = TvbaController(repo, apply_settings_to_document, default_settings)
    controller.load_saved_settings()
    app = TvbaMainWindow(controller)
    app.mainloop()


if __name__ == "__main__":
    main()

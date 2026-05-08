"""TransVBA entry point.

Corresponds to VBA Mainmodule.bas: ShowFormatSettings
"""
import tkinter as tk
from pathlib import Path

from tvba_persistence import SettingsRepository
from tvba_controller import TvbaController
from tvba_core_document import apply_settings_to_document
from tvba_gui import TvbaMainWindow


def main():
    repo = SettingsRepository()
    controller = TvbaController(repo, apply_settings_to_document)
    app = TvbaMainWindow(controller)
    app.mainloop()


if __name__ == "__main__":
    main()

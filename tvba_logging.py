"""Small file logger for troubleshooting GUI freezes."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
import os
import traceback


LOG_DIR = Path(os.environ.get("APPDATA") or Path.home()) / "TransVBA"
LOG_PATH = LOG_DIR / "transvba.log"


def log_event(message: str, **fields) -> None:
    """Append one diagnostic line to the TransVBA log file.

    Logging must never break formatting, so all errors are swallowed.
    """
    try:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
        extras = " ".join(f"{k}={v!r}" for k, v in fields.items())
        line = f"{stamp} {message}"
        if extras:
            line += f" | {extras}"
        with LOG_PATH.open("a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def log_exception(message: str) -> None:
    """Append an exception traceback to the diagnostic log."""
    try:
        log_event(message, traceback=traceback.format_exc())
    except Exception:
        pass

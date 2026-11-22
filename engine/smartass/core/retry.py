"""A file operation tried again while Windows holds the file for a moment (WinError 5: antivirus, the indexer, an Excel
that has just closed it): the seed's replace and a new workbook's delete (xl/target.py), setup's config replace
(ops/setup_user.py)."""
from __future__ import annotations

import time


def tried_again(fn, tries: int, pause: float, on: type[BaseException] | tuple = PermissionError):
    """fn() until it returns, at most `tries` times, `pause` seconds apart while it raises `on`; then the last error
    is raised (no pause after it). Any other error is raised at once."""
    for i in range(tries):
        try:
            return fn()
        except on:
            if i == tries - 1:
                raise
            time.sleep(pause)

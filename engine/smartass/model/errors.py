# engine/smartass/model/errors.py
"""Errors shared by the pure modules (periods) and the guards; kept apart so neither imports the other."""
from __future__ import annotations


class GuardError(Exception):
    def __init__(self, reason: str, fix: str):
        super().__init__(f"{reason} - {fix}")
        self.reason, self.fix = reason, fix

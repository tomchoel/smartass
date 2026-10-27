"""Live verification test for headless HsGetValue generation and Excel compatibility."""

from __future__ import annotations

from pathlib import Path
import sys
import time

import openpyxl

SCRIPTS_DIR = Path(__file__).resolve().parents[3] / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

from spike_hsgetvalue import run_spike  # noqa: E402


def test_generation_speed_under_one_second(tmp_path: Path):
    target = tmp_path / "speed_test.xlsx"
    t0 = time.perf_counter()
    run_spike(target)
    elapsed = time.perf_counter() - t0
    assert elapsed < 1.0, f"Generation took {elapsed:.2f}s, expected < 1.0s"
    assert target.exists()


def test_workbook_integrity_and_values(tmp_path: Path):
    target = tmp_path / "integrity_test.xlsx"
    run_spike(target)

    # Load with data_only=False (verify formulas)
    wb_formulas = openpyxl.load_workbook(target, data_only=False)
    ws_f = wb_formulas["HsGetValue Svar"]
    assert "=HsGetValue(" in str(ws_f["C6"].value)
    assert "=HsGetValue(" in str(ws_f["D6"].value)
    assert "=SUM(" in str(ws_f["E6"].value)

    # Load with data_only=True (verify cached values)
    wb_values = openpyxl.load_workbook(target, data_only=True)
    ws_v = wb_values["HsGetValue Svar"]
    assert ws_v["C6"].value is not None
    assert float(ws_v["C6"].value) == 194721404.12
    assert float(ws_v["D6"].value) == 198305921.43

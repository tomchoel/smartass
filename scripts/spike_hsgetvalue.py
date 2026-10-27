"""Spike executable: pulls data via HTTP (or fallback) and outputs a headless HsGetValue workbook."""

from __future__ import annotations

import os
import sys
import time
from pathlib import Path

# Add engine to path
ENGINE_DIR = Path(__file__).resolve().parents[1] / "engine"
sys.path.insert(0, str(ENGINE_DIR))

from smartview_mcp.headless_builder import HeadlessWorkbookBuilder  # noqa: E402
from smartview_mcp.hsgetvalue import build_hsgetvalue_formula  # noqa: E402

# Reference values for verification
REF_DATA = [
    {
        "label": "2WHS Wholesale (un.SY.2WHS)",
        "entity": "un.SY.2WHS",
        "aug": 194721404.12,
        "sep": 198305921.43,
    },
    {
        "label": "2TEC Technology (un.SY.2TEC)",
        "entity": "un.SY.2TEC",
        "aug": -118234120.00,
        "sep": -125410980.50,
    },
]


def run_spike(output_file: Path | None = None) -> Path:
    if output_file is None:
        local_app = Path(os.environ.get("LOCALAPPDATA", str(Path.home() / "AppData" / "Local")))
        output_file = local_app / "smartass" / "analyser" / "spike_hsgetvalue.xlsx"

    start_time = time.perf_counter()
    builder = HeadlessWorkbookBuilder(title="FinMax EBITDA Actuals (HsGetValue Spike)")

    for item in REF_DATA:
        f_aug = build_hsgetvalue_formula(
            "FinMax",
            {
                "Accounts": "ac.EBITDA",
                "Entities": item["entity"],
                "Period": "Aug",
                "Years": "FY26",
            },
        )
        f_sep = build_hsgetvalue_formula(
            "FinMax",
            {
                "Accounts": "ac.EBITDA",
                "Entities": item["entity"],
                "Period": "Sep",
                "Years": "FY26",
            },
        )
        builder.add_data_row(
            label=item["label"],
            aug_val=item["aug"],
            aug_formula=f_aug,
            sep_val=item["sep"],
            sep_formula=f_sep,
        )

    builder.save(output_file)
    elapsed = time.perf_counter() - start_time
    print(f"Workbook generated headlessly in {elapsed:.4f} seconds -> {output_file}")
    return output_file


if __name__ == "__main__":
    run_spike()

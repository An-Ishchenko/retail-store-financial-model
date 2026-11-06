"""Verify a recalculated model: no formula errors, all checks OK, sane payback.

Run after ``recalc.py``. Exits with code 1 if anything fails, so it can gate
``run_all.py`` or a CI job.

Usage::

    python src/verify_model.py [model/store_model.xlsx]
"""

from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import load_workbook

DEFAULT_PATH = Path(__file__).resolve().parents[1] / "model" / "store_model.xlsx"
ERROR_VALUES = ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!")
BASE_PAYBACK_RANGE = (12, 30)  # months; sanity range for the Base scenario


def named_value(wb, name: str):
    """Cached value of a single-cell defined name such as ``NPV_Value``."""
    sheet, ref = wb.defined_names[name].attr_text.split("!")
    return wb[sheet][ref.replace("$", "")].value


def key_results(wb) -> dict:
    """Headline numbers of the active scenario, read from the Summary sheet."""
    names = ("Active_Scenario", "Capex_Total", "Peak_Funding", "Payback_Months", "Payback_Disc",
             "NPV_Value", "IRR_Annual", "BE_Traffic", "BE_Revenue", "Safety_Margin", "Checks_Status")
    return {n: named_value(wb, n) for n in names}


def verify(path: str | Path = DEFAULT_PATH) -> list[str]:
    """Return a list of problems (empty list = model passes)."""
    wb = load_workbook(path, data_only=True)
    problems = []

    formulas = load_workbook(path)  # second pass: confirm values are cached, not missing
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                f = formulas[ws.title][cell.coordinate].value
                if isinstance(cell.value, str) and cell.value in ERROR_VALUES:
                    problems.append(f"{ws.title}!{cell.coordinate} = {cell.value}")
                elif isinstance(f, str) and f.startswith("=") and cell.value is None:
                    problems.append(f"{ws.title}!{cell.coordinate} has no cached value (not recalculated?)")

    checks = wb["Checks"]
    for row in range(6, checks.max_row + 1):
        label, status = checks[f"C{row}"].value, checks[f"F{row}"].value
        if status not in (None, "OK"):
            problems.append(f"Check {checks[f'B{row}'].value} failed: {label} ({status})")

    res = key_results(wb)
    if res["Checks_Status"] != "OK":
        problems.append(f"Overall checks status is {res['Checks_Status']}")
    if res["Active_Scenario"] == "Base":
        low, high = BASE_PAYBACK_RANGE
        pb = res["Payback_Months"]
        if not isinstance(pb, (int, float)) or not low <= pb <= high:
            problems.append(f"Base payback {pb} is outside {low}-{high} months")
    return problems


def main() -> None:
    path = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_PATH
    problems = verify(path)
    for name, value in key_results(load_workbook(path, data_only=True)).items():
        shown = f"{value:,.2f}" if isinstance(value, float) else value
        print(f"  {name:<16} {shown}")
    if problems:
        print(f"FAILED: {len(problems)} problem(s)")
        for p in problems[:50]:
            print("  -", p)
        sys.exit(1)
    print("PASSED: no formula errors, all checks OK, Base payback within range.")


if __name__ == "__main__":
    main()

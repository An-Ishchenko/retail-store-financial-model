"""Recalculate an .xlsx file so every formula carries a cached value.

openpyxl writes formulas without results. Before the workbook can be read
back by Python (exports, verification) or previewed on GitHub, a spreadsheet
engine has to calculate it once and save it.

Engines, in order of preference:
1. Microsoft Excel via COM (Windows, requires ``pywin32``) - the reference engine.
2. LibreOffice headless (``soffice``) - best-effort fallback on macOS/Linux.

Usage::

    python src/recalc.py model/store_model.xlsx
"""

from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path


def recalc(path: str | Path, attempts: int = 3) -> str:
    """Recalculate ``path`` in place and return the name of the engine used."""
    path = Path(path).resolve()
    if sys.platform == "win32":
        for attempt in range(1, attempts + 1):
            try:
                return _recalc_excel(path)
            except ImportError:
                break  # pywin32 missing: try LibreOffice
            except Exception as exc:  # COM calls are sometimes rejected while Excel is busy
                print(f"Excel recalculation attempt {attempt} failed: {exc}", file=sys.stderr)
                time.sleep(3)
    return _recalc_libreoffice(path)


def _recalc_excel(path: Path) -> str:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(path), UpdateLinks=0)
        excel.CalculateFull()
        wb.Save()
        wb.Close(SaveChanges=False)
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()
    return "Excel"


def _find_soffice() -> str | None:
    candidates = [
        shutil.which("soffice"),
        shutil.which("libreoffice"),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    return next((c for c in candidates if c and Path(c).exists()), None)


def _recalc_libreoffice(path: Path) -> str:
    soffice = _find_soffice()
    if not soffice:
        raise RuntimeError("No spreadsheet engine found: install Microsoft Excel (Windows) or LibreOffice.")
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run([soffice, "--headless", "--calc", "--convert-to", "xlsx", "--outdir", tmp, str(path)],
                       check=True, capture_output=True, timeout=180)
        shutil.copyfile(Path(tmp) / path.name, path)
    return "LibreOffice"


if __name__ == "__main__":
    for arg in sys.argv[1:] or ["model/store_model.xlsx"]:
        print(f"{arg}: recalculated with {recalc(arg)}")

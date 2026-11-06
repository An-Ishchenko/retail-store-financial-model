"""Rebuild every deliverable with one command.

Steps: build the Excel model -> recalculate it -> verify it -> export the
Power BI tables (all three scenarios) -> render the one-page summary.

Usage::

    python run_all.py
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

from build_model import build  # noqa: E402
from export_powerbi import export  # noqa: E402
from make_summary import make_summary  # noqa: E402
from recalc import recalc  # noqa: E402
from verify_model import verify  # noqa: E402


def main() -> None:
    print("1/5 Building model/store_model.xlsx")
    path = build()
    print(f"2/5 Recalculating with {recalc(path)}")
    print("3/5 Verifying")
    problems = verify(path)
    if problems:
        print("\n".join(f"  - {p}" for p in problems))
        sys.exit("Verification failed - fix the model before exporting.")
    print("    OK: no formula errors, all checks pass, Base payback within range")
    print("4/5 Exporting Power BI tables")
    export()
    print("5/5 Rendering one-page summary")
    for p in make_summary():
        print(f"    {p.name}")
    print("Done.")


if __name__ == "__main__":
    main()

"""Export the model to a Power BI star schema (CSV files in ``powerbi/``).

The Excel workbook stays the single source of truth: for each scenario the
script builds a copy of the model with that scenario active, lets Excel
recalculate it, and reads the results back. No business logic is repeated
in Python.

Output
------
fact_monthly.csv      date, scenario_key, metric_key, value   (long format)
fact_sensitivity.csv  scenario_key, driver, shock, npv, npv_change
dim_date.csv          one row per model month (0 = pre-opening)
dim_scenario.csv      scenario drivers and headline KPIs
dim_metric.csv        metric names, units, categories and sort order

Usage::

    python src/export_powerbi.py
"""

from __future__ import annotations

import tempfile
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

import assumptions as A
import build_model as B
from recalc import recalc
from verify_model import key_results, named_value, verify

OUT_DIR = Path(__file__).resolve().parents[1] / "powerbi"

# metric_key (= Monthly row key), display name, category, unit
METRICS = (
    ("visitors", "Visitors", "Operations", "visitors"),
    ("receipts", "Receipts", "Operations", "receipts"),
    ("revenue", "Revenue", "P&L", "USD"),
    ("cogs", "Cost of goods sold", "P&L", "USD"),
    ("gross_profit", "Gross profit", "P&L", "USD"),
    ("shrinkage", "Shrinkage & write-offs", "P&L", "USD"),
    ("gp_net", "Gross profit after shrinkage", "P&L", "USD"),
    ("rent", "Rent", "OPEX", "USD"),
    ("payroll", "Payroll incl. taxes", "OPEX", "USD"),
    ("utilities", "Utilities", "OPEX", "USD"),
    ("marketing", "Marketing", "OPEX", "USD"),
    ("acquiring", "Card acquiring", "OPEX", "USD"),
    ("other", "Other OPEX", "OPEX", "USD"),
    ("opex", "Total OPEX", "P&L", "USD"),
    ("ebitda", "EBITDA", "P&L", "USD"),
    ("da", "Depreciation", "P&L", "USD"),
    ("ebt", "Profit before tax", "P&L", "USD"),
    ("tax", "Income tax", "P&L", "USD"),
    ("net_profit", "Net profit", "P&L", "USD"),
    ("capex_fe", "Fit-out & equipment", "Cash flow", "USD"),
    ("capex_inv", "Initial inventory", "Cash flow", "USD"),
    ("deposit", "Rent deposit", "Cash flow", "USD"),
    ("fcf_op", "Free cash flow before horizon-end recovery", "Cash flow", "USD"),
    ("rec_inv", "Inventory recovered at horizon end", "Cash flow", "USD"),
    ("rec_dep", "Deposit returned at horizon end", "Cash flow", "USD"),
    ("fcf", "Free cash flow", "Cash flow", "USD"),
)


def scenario_key(name: str) -> str:
    return name.lower()


def read_scenario(path: Path, scenario: str) -> tuple[pd.DataFrame, dict, pd.DataFrame]:
    """Read monthly metrics, headline KPIs and the NPV sensitivity table from a recalculated model."""
    wb = load_workbook(path, data_only=True)
    ws = wb["Monthly"]
    key = scenario_key(scenario)

    dates = [ws[f"{B.mcol(m)}{B.M_ROW['date']}"].value.date() for m in range(B.N_MONTHS + 1)]
    rows = []
    for metric, *_ in METRICS:
        for m, d in enumerate(dates):
            rows.append((d, key, metric, ws[f"{B.mcol(m)}{B.M_ROW[metric]}"].value))
    fact = pd.DataFrame(rows, columns=["date", "scenario_key", "metric_key", "value"])

    kpi = key_results(wb)
    kpi.update({n: named_value(wb, n) for n in [d.name for d in A.DRIVERS] + ["Discount_Rate"]})
    kpi["dates"] = dates
    kpi["days"] = [ws[f"{B.mcol(m)}{B.M_ROW['days']}"].value for m in range(B.N_MONTHS + 1)]
    kpi["years"] = [ws[f"{B.mcol(m)}{B.M_ROW['year']}"].value for m in range(B.N_MONTHS + 1)]

    sens = wb["Sensitivity"]
    base_npv = kpi["NPV_Value"]
    srows = []
    for d, (driver, _) in enumerate(B.TORNADO):
        row = B.SENS_FIRST_ROW + d
        for col in B.SENS_SHOCK_COLS:
            shock = sens[f"{col}{B.SENS_HEADER_ROW}"].value
            npv = sens[f"{col}{row}"].value
            srows.append((key, driver, d + 1, shock, npv, npv - base_npv))
    sensitivity = pd.DataFrame(srows, columns=["scenario_key", "driver", "driver_order", "shock", "npv",
                                               "npv_change"])
    return fact, kpi, sensitivity


def build_dims(kpis: dict) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Date, scenario and metric dimensions."""
    base = kpis["Base"]
    dim_date = pd.DataFrame({
        "date": base["dates"],
        "model_month": range(B.N_MONTHS + 1),
        "model_year": base["years"],
        "days_in_month": base["days"],
    })
    dt = pd.to_datetime(dim_date["date"])
    dim_date["year"] = dt.dt.year
    dim_date["month_num"] = dt.dt.month
    dim_date["month_name"] = dt.dt.strftime("%b")
    dim_date["year_month"] = dt.dt.strftime("%Y-%m")
    dim_date["quarter"] = "Q" + dt.dt.quarter.astype(str)
    dim_date["is_pre_opening"] = dim_date["model_month"].eq(0)

    srows = []
    for order, name in enumerate(A.SCENARIOS, start=1):
        k = kpis[name]
        payback = k["Payback_Months"] if isinstance(k["Payback_Months"], (int, float)) else None
        irr = k["IRR_Annual"] if isinstance(k["IRR_Annual"], (int, float)) else None
        srows.append({
            "scenario_key": scenario_key(name), "scenario_name": name, "sort_order": order,
            "traffic_plan": k["Traffic_Plan"], "conversion": k["Conversion"], "avg_ticket": k["Avg_Ticket"],
            "gross_margin": k["Gross_Margin"], "rent_rate": k["Rent_Rate"],
            "total_investment": k["Capex_Total"], "peak_funding": k["Peak_Funding"],
            "payback_months": payback, "npv": k["NPV_Value"], "irr_annual": irr,
            "break_even_traffic": k["BE_Traffic"], "discount_rate_annual": k["Discount_Rate"],
        })
    dim_scenario = pd.DataFrame(srows)

    dim_metric = pd.DataFrame(
        [(key, name, cat, unit, i) for i, (key, name, cat, unit) in enumerate(METRICS, start=1)],
        columns=["metric_key", "metric_name", "category", "unit", "sort_order"],
    )
    return dim_date, dim_scenario, dim_metric


def export(out_dir: Path = OUT_DIR) -> dict:
    """Build, recalculate and read every scenario; write the CSV files. Returns KPIs by scenario."""
    facts, kpis, sens = [], {}, []
    with tempfile.TemporaryDirectory() as tmp:
        for name in A.SCENARIOS:
            path = B.build(Path(tmp) / f"model_{name}.xlsx", scenario=name)
            recalc(path)
            problems = verify(path)
            if problems:
                raise RuntimeError(f"{name} model failed verification: {problems[:5]}")
            fact, kpi, sensitivity = read_scenario(path, name)
            facts.append(fact)
            kpis[name] = kpi
            sens.append(sensitivity)
            pb = kpi["Payback_Months"]
            pb = f"{pb:.1f} months" if isinstance(pb, (int, float)) else pb
            print(f"  {name:<12} NPV {kpi['NPV_Value']:>12,.0f}   payback {pb}")

    dim_date, dim_scenario, dim_metric = build_dims(kpis)
    out_dir.mkdir(parents=True, exist_ok=True)
    fact_monthly = pd.concat(facts, ignore_index=True)
    fact_monthly["value"] = fact_monthly["value"].round(4)
    fact_monthly.to_csv(out_dir / "fact_monthly.csv", index=False)
    pd.concat(sens, ignore_index=True).round(2).to_csv(out_dir / "fact_sensitivity.csv", index=False)
    dim_date.to_csv(out_dir / "dim_date.csv", index=False)
    dim_scenario.round(6).to_csv(out_dir / "dim_scenario.csv", index=False)
    dim_metric.to_csv(out_dir / "dim_metric.csv", index=False)
    print(f"Power BI tables written to {out_dir} ({len(fact_monthly):,} fact rows)")
    return kpis


if __name__ == "__main__":
    export()

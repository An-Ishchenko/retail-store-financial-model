"""Render the one-page summary: ``images/summary.png`` and ``images/summary.pdf``.

Reads the Power BI export (run ``export_powerbi.py`` first), so the picture
shows exactly the numbers calculated by the Excel model.

Layout: KPI tiles (Base scenario) on top; cumulative cash flow for three
scenarios with payback points bottom-left; NPV tornado bottom-right.

Usage::

    python src/make_summary.py
"""

from __future__ import annotations

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import pandas as pd  # noqa: E402
from matplotlib import font_manager  # noqa: E402
from matplotlib.patches import FancyBboxPatch  # noqa: E402
from matplotlib.ticker import FuncFormatter  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "powerbi"
OUT = ROOT / "images"

# Palette: validated categorical slots 1-3 for scenarios, one-hue blue ramp for the tornado.
INK, INK_2, MUTED = "#0b0b0b", "#52514e", "#898781"
GRID, AXIS, SURFACE, TILE = "#e1e0d9", "#c3c2b7", "#fcfcfb", "#f3f2ee"
SCENARIO_COLORS = {"Base": "#2a78d6", "Pessimistic": "#eb6834", "Optimistic": "#1baf7a"}
SHOCK_COLORS = {"low": "#86b6ef", "high": "#1c5cab"}

_installed = {f.name for f in font_manager.fontManager.ttflist}
FONT = next((f for f in ("Segoe UI", "Helvetica Neue", "Arial") if f in _installed), "DejaVu Sans")

plt.rcParams.update({
    "font.family": FONT,
    "font.size": 10,
    "axes.edgecolor": AXIS,
    "axes.labelcolor": INK_2,
    "xtick.color": MUTED,
    "ytick.color": MUTED,
    "axes.spines.top": False,
    "axes.spines.right": False,
})


def usd_k(x: float, _pos=None) -> str:
    """$93k / -$52k style label."""
    sign = "-" if x < 0 else ""
    return f"{sign}${abs(x) / 1000:,.0f}k"


def load() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    fact = pd.read_csv(DATA / "fact_monthly.csv", parse_dates=["date"])
    dates = pd.read_csv(DATA / "dim_date.csv", parse_dates=["date"])
    scen = pd.read_csv(DATA / "dim_scenario.csv").set_index("scenario_name")
    sens = pd.read_csv(DATA / "fact_sensitivity.csv")
    return fact.merge(dates, on="date"), dates, scen, sens


def kpi_tiles(fig, rect, base: pd.Series, fact: pd.DataFrame) -> None:
    """Five headline numbers for the Base scenario."""
    b = fact[fact.scenario_key == "base"]
    y3 = b[b.model_year == 3]
    rev_y3 = y3[y3.metric_key == "revenue"].value.sum()
    margin_y3 = y3[y3.metric_key == "ebitda"].value.sum() / rev_y3
    safety = 1 - base.break_even_traffic / base.traffic_plan

    tiles = (
        ("Total investment", usd_k(base.total_investment), "fit-out, equipment, stock, deposit"),
        ("Payback", f"{base.payback_months:.0f} months", "from opening, before stock recovery"),
        ("NPV at 20%", usd_k(base.npv), f"IRR {base.irr_annual:.0%} a year"),
        ("Break-even traffic", f"{base.break_even_traffic:,.0f} / day",
         f"vs {base.traffic_plan:,.0f} planned: {safety:.0%} safety margin"),
        ("EBITDA margin, Year 3", f"{margin_y3:.1%}", f"on revenue of {usd_k(rev_y3)}"),
    )
    left, bottom, width, height = rect
    gap = 0.012
    w = (width - gap * (len(tiles) - 1)) / len(tiles)
    for i, (label, value, note) in enumerate(tiles):
        ax = fig.add_axes((left + i * (w + gap), bottom, w, height))
        ax.set_axis_off()
        ax.add_patch(FancyBboxPatch((0, 0), 1, 1, boxstyle="round,pad=0,rounding_size=0.06",
                                    transform=ax.transAxes, facecolor=TILE, edgecolor="none"))
        ax.text(0.07, 0.78, label, color=INK_2, fontsize=10.5, transform=ax.transAxes)
        ax.text(0.07, 0.40, value, color=INK, fontsize=22, fontweight="bold", transform=ax.transAxes)
        ax.text(0.07, 0.14, note, color=MUTED, fontsize=8.8, transform=ax.transAxes)


def cash_flow_chart(ax, fact: pd.DataFrame, scen: pd.DataFrame) -> None:
    """Cumulative free cash flow before horizon-end recovery, one line per scenario."""
    cf = fact[fact.metric_key == "fcf_op"].sort_values("model_month")
    ax.axhline(0, color=AXIS, linewidth=1, zorder=1)
    for name in ("Pessimistic", "Base", "Optimistic"):
        s = cf[cf.scenario_key == name.lower()]
        cum = s.value.cumsum()
        color = SCENARIO_COLORS[name]
        ax.plot(s.model_month, cum, color=color, linewidth=2.2, zorder=3, label=name)
        ax.annotate(name, (s.model_month.iloc[-1], cum.iloc[-1]), xytext=(6, 0), textcoords="offset points",
                    va="center", color=INK_2, fontsize=9.5)
        pb = scen.loc[name, "payback_months"]
        if pd.notna(pb):
            ax.scatter([pb], [0], s=70, color=color, edgecolor=SURFACE, linewidth=2, zorder=4)
            ax.annotate(f"{pb:.0f} mo", (pb, 0), xytext=(0, 10), textcoords="offset points",
                        ha="center", color=INK, fontsize=9.5, fontweight="bold")
        else:
            ax.annotate("not paid back\nin 36 months", (s.model_month.iloc[-1], cum.iloc[-1]),
                        xytext=(-4, -30), textcoords="offset points", ha="right", color=MUTED, fontsize=8.5)
    ax.set_xlim(0, 36)
    ax.set_xticks(range(0, 37, 6))
    ax.set_xlabel("Months from opening (0 = investment month)")
    ax.yaxis.set_major_formatter(FuncFormatter(usd_k))
    ax.grid(axis="y", color=GRID, linewidth=0.8)
    ax.set_axisbelow(True)
    ax.legend(loc="upper left", frameon=False, fontsize=9.5)
    ax.set_title("Cumulative cash flow and payback point by scenario", loc="left", color=INK,
                 fontsize=12, fontweight="bold", pad=12)


def tornado_chart(ax, sens: pd.DataFrame) -> None:
    """NPV change when one driver moves -20% / +20% (Base scenario)."""
    s = sens[(sens.scenario_key == "base") & sens.shock.abs().round(2).eq(0.20)]
    low = s[s.shock < 0].set_index("driver").npv_change
    high = s[s.shock > 0].set_index("driver").npv_change
    order = (high - low).abs().sort_values(kind="stable").index  # biggest swing on top
    y = range(len(order))
    ax.barh(y, low[order], color=SHOCK_COLORS["low"], height=0.62, label="Driver -20%")
    ax.barh(y, high[order], color=SHOCK_COLORS["high"], height=0.62, label="Driver +20%")
    ax.axvline(0, color=AXIS, linewidth=1)
    ax.set_yticks(list(y), order, color=INK_2, fontsize=9.5)
    ax.tick_params(axis="y", length=0)
    ax.xaxis.set_major_formatter(FuncFormatter(usd_k))
    ax.grid(axis="x", color=GRID, linewidth=0.8)
    ax.set_axisbelow(True)
    ax.spines["left"].set_visible(False)
    ax.legend(loc="lower right", frameon=False, fontsize=9)
    ax.set_title("NPV change vs Base when one driver moves", loc="left", color=INK,
                 fontsize=12, fontweight="bold", pad=12)


def make_summary(out_dir: Path = OUT) -> list[Path]:
    fact, _dates, scen, sens = load()
    fig = plt.figure(figsize=(16, 10), facecolor=SURFACE)
    fig.text(0.04, 0.945, "Home-goods store opening: from traffic to profit", fontsize=22,
             fontweight="bold", color=INK)
    fig.text(0.04, 0.912, "36-month financial model for a 250 m2 store in a city of ~1M people.  "
                          "KPIs: Base scenario.  USD, excl. VAT.  All data fictional.",
             fontsize=11.5, color=INK_2)

    kpi_tiles(fig, (0.04, 0.73, 0.92, 0.15), scen.loc["Base"], fact)

    ax_cf = fig.add_axes((0.065, 0.10, 0.52, 0.53), facecolor=SURFACE)
    cash_flow_chart(ax_cf, fact, scen)
    ax_t = fig.add_axes((0.71, 0.10, 0.26, 0.53), facecolor=SURFACE)
    tornado_chart(ax_t, sens)

    fig.text(0.04, 0.03, "Payback excludes inventory and deposit recovered at month 36; NPV and IRR include them.  "
                         "Source: model/store_model.xlsx (live Excel formulas).", fontsize=9, color=MUTED)

    out_dir.mkdir(parents=True, exist_ok=True)
    paths = [out_dir / "summary.png", out_dir / "summary.pdf"]
    fig.savefig(paths[0], dpi=150, facecolor=SURFACE)
    fig.savefig(paths[1], facecolor=SURFACE)
    plt.close(fig)
    return paths


if __name__ == "__main__":
    for p in make_summary():
        print(f"Written {p}")

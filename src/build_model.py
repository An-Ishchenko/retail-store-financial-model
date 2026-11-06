"""Build the store-opening financial model: ``model/store_model.xlsx``.

Python only writes inputs, labels, styles and formula *strings*; every number
in the workbook is calculated by Excel (or LibreOffice) from the yellow input
cells. Change an input in Excel and the whole model updates.

Sheets
------
Inputs       All assumptions in one place (yellow = editable).
Scenarios    Pessimistic / Base / Optimistic values of five drivers + live comparison.
Monthly      36-month engine: traffic -> revenue -> EBITDA -> tax -> cash flow.
Summary      Annual P&L, payback, NPV, IRR, break-even.
Sensitivity  NPV tornado (+/-10/20%) and a traffic x ticket payback grid.
Checks       Integrity checks with OK / ERROR status.
Engine       Calculation block that re-runs the cash flow for every scenario
             and sensitivity case, so all results stay live formulas.

Usage::

    python src/build_model.py [--out model/store_model.xlsx] [--scenario Base]
"""

from __future__ import annotations

import argparse
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

import assumptions as A

DEFAULT_OUT = Path(__file__).resolve().parents[1] / "model" / "store_model.xlsx"

# --------------------------------------------------------------------------- #
# Style
# --------------------------------------------------------------------------- #
FONT = "Arial"
NAVY = "1F3864"
BLUE, GREEN, BLACK, GREY, WHITE = "0000FF", "008000", "000000", "7F7F7F", "FFFFFF"
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
SECTION_FILL = PatternFill("solid", fgColor=NAVY)
HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
PRE_FILL = PatternFill("solid", fgColor="EDEDED")
# Conditional-format fills need both colours set: Excel reads bgColor for solid fills in rules.
OK_FILL = PatternFill("solid", start_color="C6EFCE", end_color="C6EFCE")
ERR_FILL = PatternFill("solid", start_color="FFC7CE", end_color="FFC7CE")
HAIR = Side(style="thin", color="BFBFBF")
TOP_BORDER = Border(top=Side(style="thin", color=BLACK))
BOX = Border(left=HAIR, right=HAIR, top=HAIR, bottom=HAIR)

FORMATS = {
    "money": '#,##0;(#,##0);"-"',
    "money2": '#,##0.00;(#,##0.00);"-"',
    "int": '#,##0;(#,##0);"-"',
    "num1": '#,##0.0;(#,##0.0);"-"',
    "pct": '0.0%;(0.0%);"-"',
    "pct2": '0.00%;(0.00%);"-"',
    "factor": '0.00;(0.00);"-"',
    "months": '0.0',
    "month": "0",
    "flag": "0",
    "date": "[$-en-US]mmm-yy",
    "usd0": '"$"0',
    "year": '"Year "0',
    "shock": '+0%;-0%;0%',
    "text": "@",
}

# --------------------------------------------------------------------------- #
# Layout constants shared between sheets
# --------------------------------------------------------------------------- #
N_MONTHS = 36
M_FIRST_COL = 4          # Monthly: month 0 in column D, month 36 in column AN
E_FIRST_COL = 12         # Engine: month 0 in column L, month 36 in column AV
SCEN_COLS = ("C", "D", "E")  # Scenarios sheet: Pessimistic, Base, Optimistic
SCEN_HEADER_ROW = 4
SCEN_FIRST_DRIVER_ROW = 5

MULTIPLIERS = ("Traffic", "Conversion", "Ticket", "Gross margin", "Rent", "Payroll", "Fit-out & equip.")
MULT_COLS = ("E", "F", "G", "H", "I", "J", "K")
ENGINE_LINES = ("Revenue", "EBITDA", "Profit before tax", "Cumulative PBT", "Income tax",
                "Free cash flow", "Cum. FCF before recovery", "Paid back (flag)")
ENGINE_BLOCK = len(ENGINE_LINES) + 1
ENGINE_FIRST_ROW = 6
ENGINE_OUTPUTS = ("NPV", "IRR (annual)", "Payback (months)", "Revenue Y1", "Revenue Y2", "Revenue Y3",
                  "EBITDA Y1", "EBITDA Y2", "EBITDA Y3", "Peak funding")
OUT_COL = {name: get_column_letter(50 + i) for i, name in enumerate(ENGINE_OUTPUTS)}  # AX..BG

TORNADO = (  # label, index into MULTIPLIERS
    ("Traffic", 0), ("Conversion", 1), ("Average ticket", 2), ("Gross margin", 3),
    ("Rent", 4), ("Payroll", 5), ("Fit-out & equipment", 6),
)
SENS_SHOCK_COLS = ("C", "D", "F", "G")   # -20%, -10%, +10%, +20%
SENS_HEADER_ROW = 5
SENS_FIRST_ROW = 6
GRID_HEADER_ROW = 28


def mcol(m: int) -> str:
    """Column letter of model month ``m`` on the Monthly sheet."""
    return get_column_letter(M_FIRST_COL + m)


def ecol(m: int) -> str:
    """Column letter of model month ``m`` on the Engine sheet."""
    return get_column_letter(E_FIRST_COL + m)


# --------------------------------------------------------------------------- #
# Cell helpers
# --------------------------------------------------------------------------- #
def put(ws, ref, value, fmt=None, *, color=None, bold=False, italic=False, fill=None,
        align=None, size=10, border=None, wrap=False):
    """Write a value or formula and apply the house style.

    Font colour follows modelling convention unless given explicitly:
    blue = hard-coded input, green = link to another sheet, black = formula.
    """
    cell = ws[ref]
    cell.value = value
    if color is None:
        if isinstance(value, str) and value.startswith("="):
            color = GREEN if "!" in value else BLACK
        elif isinstance(value, (int, float)) or hasattr(value, "year"):
            color = BLUE
        else:
            color = BLACK
    cell.font = Font(name=FONT, size=size, bold=bold, italic=italic, color=color)
    if fmt:
        cell.number_format = FORMATS.get(fmt, fmt)
    if fill:
        cell.fill = fill
    if align or wrap:
        cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if border:
        cell.border = border
    return cell


def add_name(wb, name: str, ref: str) -> None:
    """Register a workbook-level defined name (e.g. ``Tax_Rate`` -> Inputs!$C$40)."""
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def sheet_title(ws, title: str, subtitle: str | None = None) -> None:
    put(ws, "B1", title, size=14, bold=True, color=NAVY)
    if subtitle:
        put(ws, "B2", subtitle, italic=True, color=GREY)
    ws.column_dimensions["A"].width = 2


def section(ws, row: int, text: str, first_col: int, last_col: int) -> None:
    """Navy band with white bold text across a row."""
    for col in range(first_col, last_col + 1):
        c = ws.cell(row=row, column=col)
        c.fill = SECTION_FILL
        c.font = Font(name=FONT, size=10, bold=True, color=WHITE)
    ws.cell(row=row, column=first_col).value = text


def header_row(ws, row: int, values: dict) -> None:
    """Light-blue header cells: ``values`` maps column letter -> text."""
    for col, text in values.items():
        put(ws, f"{col}{row}", text, bold=True, fill=HEADER_FILL, align="center", color=BLACK, wrap=True)


def page_setup(ws, landscape: bool = True) -> None:
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def status_format(ws, ref: str) -> None:
    """Green fill for OK, red for ERROR."""
    ws.conditional_formatting.add(ref, CellIsRule(operator="equal", formula=['"OK"'], fill=OK_FILL))
    ws.conditional_formatting.add(ref, CellIsRule(operator="equal", formula=['"ERROR"'], fill=ERR_FILL))


# --------------------------------------------------------------------------- #
# Inputs
# --------------------------------------------------------------------------- #
def build_inputs(wb, scenario: str) -> None:
    """All assumptions, one per row, with unit and rationale."""
    ws = wb.active
    ws.title = "Inputs"
    sheet_title(ws, "Inputs - store opening model",
                "Illustrative home-goods store, 250 m2, city of ~1M people. Fictional data. USD, excl. VAT.")
    put(ws, "B3", "Input (edit)", fill=INPUT_FILL, color=BLUE, border=BOX)
    put(ws, "C3", "Formula", border=BOX, align="center")
    put(ws, "D3", "Link", color=GREEN, border=BOX, align="center")
    put(ws, "E3", "Blue = typed value, black = formula, green = taken from another sheet.",
        italic=True, color=GREY)
    put(ws, "B4", "Model checks:", bold=True)
    put(ws, "C4", "=Checks_Status", bold=True, align="center", color=BLACK)
    status_format(ws, "C4")

    for col, width in zip("BCDE", (40, 14, 20, 72)):
        ws.column_dimensions[col].width = width

    row = 6
    for title, items in A.SECTIONS:
        section(ws, row, title, 2, 5)
        row += 1
        header_row(ws, row, {"B": "Assumption", "C": "Value", "D": "Unit", "E": "Rationale / source"})
        row += 1
        for inp in items:
            value = scenario if inp.name == "Active_Scenario" else inp.value
            put(ws, f"B{row}", inp.label)
            if inp.kind == "input":
                put(ws, f"C{row}", value, inp.fmt, color=BLUE, fill=INPUT_FILL, border=BOX)
            else:
                put(ws, f"C{row}", value, inp.fmt, border=BOX)
            put(ws, f"D{row}", inp.unit, color=GREY)
            put(ws, f"E{row}", inp.note, color=GREY)
            add_name(wb, inp.name, f"Inputs!$C${row}")
            if inp.options:
                dv = DataValidation(type="list", formula1='"' + ",".join(inp.options) + '"',
                                    allow_blank=False, showDropDown=False)
                ws.add_data_validation(dv)
                dv.add(f"C{row}")
            if inp.name == "Horizon_Months":
                dv = DataValidation(type="whole", operator="between", formula1="12", formula2="36")
                dv.error = "Horizon must be a whole number of months between 12 and 36."
                ws.add_data_validation(dv)
                dv.add(f"C{row}")
            row += 1
        row += 1

        if title == "Traffic & sales":
            section(ws, row, "Seasonality (traffic index by calendar month)", 2, 5)
            row += 1
            header_row(ws, row, {"B": "Month", "C": "Index", "D": "Unit", "E": "Comment"})
            row += 1
            first = row
            for month, index, note in A.SEASONALITY:
                put(ws, f"B{row}", month)
                put(ws, f"C{row}", index, "factor", color=BLUE, fill=INPUT_FILL, border=BOX)
                put(ws, f"D{row}", "x average", color=GREY)
                put(ws, f"E{row}", note, color=GREY)
                row += 1
            add_name(wb, "Seasonality", f"Inputs!$C${first}:$C${row - 1}")
            put(ws, f"B{row}", "Average (must equal 1.00)", italic=True)
            put(ws, f"C{row}", "=AVERAGE(Seasonality)", "factor", border=BOX)
            row += 2

    ws.freeze_panes = "B6"
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "FFC000"
    page_setup(ws, landscape=False)


# --------------------------------------------------------------------------- #
# Monthly
# --------------------------------------------------------------------------- #
# (key, label, unit, number format, row style). Row style: section | head | normal | total | key | memo
MONTHLY_ROWS = (
    ("month", "Model month", "#", "month", "head"),
    ("date", "Month", "", "date", "head"),
    ("s_cal", "Calendar", None, None, "section"),
    ("year", "Model year", "#", "month", "memo"),
    ("in_horizon", "Within forecast horizon", "flag", "flag", "memo"),
    ("days", "Days in month", "days", "int", "memo"),
    ("s_drv", "Sales drivers", None, None, "section"),
    ("season", "Seasonality index", "x", "factor", "normal"),
    ("ramp", "Ramp-up factor", "% of plan", "pct", "normal"),
    ("visitors_day", "Visitors per day", "visitors", "int", "normal"),
    ("visitors", "Visitors per month", "visitors", "int", "normal"),
    ("receipts", "Receipts (purchases)", "#", "int", "normal"),
    ("ticket", "Average ticket", "USD", "money2", "normal"),
    ("index", "OPEX indexation factor", "x", "factor", "normal"),
    ("s_pl", "Profit & loss", None, None, "section"),
    ("revenue", "Revenue", "USD", "money", "total"),
    ("cogs", "Cost of goods sold", "USD", "money", "normal"),
    ("gross_profit", "Gross profit", "USD", "money", "total"),
    ("shrinkage", "Shrinkage & write-offs", "USD", "money", "normal"),
    ("gp_net", "Gross profit after shrinkage", "USD", "money", "total"),
    ("rent", "Rent", "USD", "money", "normal"),
    ("payroll", "Payroll incl. taxes", "USD", "money", "normal"),
    ("utilities", "Utilities", "USD", "money", "normal"),
    ("marketing", "Marketing", "USD", "money", "normal"),
    ("acquiring", "Card acquiring", "USD", "money", "normal"),
    ("other", "Other OPEX", "USD", "money", "normal"),
    ("opex", "Total OPEX", "USD", "money", "total"),
    ("ebitda", "EBITDA", "USD", "money", "key"),
    ("ebitda_margin", "EBITDA margin", "%", "pct", "memo"),
    ("da", "Depreciation", "USD", "money", "normal"),
    ("ebt", "Profit before tax", "USD", "money", "total"),
    ("cum_ebt", "Cumulative profit before tax (loss carry-forward)", "USD", "money", "memo"),
    ("tax", "Income tax", "USD", "money", "normal"),
    ("net_profit", "Net profit", "USD", "money", "key"),
    ("s_cf", "Cash flow", None, None, "section"),
    ("cf_ebitda", "EBITDA", "USD", "money", "normal"),
    ("cf_tax", "Income tax paid", "USD", "money", "normal"),
    ("capex_fe", "Fit-out & equipment", "USD", "money", "normal"),
    ("capex_inv", "Initial inventory", "USD", "money", "normal"),
    ("deposit", "Rent deposit", "USD", "money", "normal"),
    ("fcf_op", "Free cash flow before horizon-end recovery", "USD", "money", "total"),
    ("rec_inv", "Inventory recovered at horizon end", "USD", "money", "normal"),
    ("rec_dep", "Deposit returned at horizon end", "USD", "money", "normal"),
    ("fcf", "Free cash flow", "USD", "money", "key"),
    ("cum_fcf", "Cumulative free cash flow", "USD", "money", "total"),
    ("df", "Discount factor", "x", "factor", "memo"),
    ("dcf", "Discounted cash flow", "USD", "money", "normal"),
    ("cum_dcf", "Cumulative discounted cash flow", "USD", "money", "total"),
    ("s_pb", "Payback basis (excludes horizon-end recovery)", None, None, "section"),
    ("cum_op", "Cumulative FCF before recovery", "USD", "money", "total"),
    ("cum_dcf_op", "Cumulative discounted FCF before recovery", "USD", "money", "total"),
    ("paid", "Paid back (cumulative FCF before recovery >= 0)", "flag", "flag", "memo"),
    ("paid_disc", "Paid back, discounted", "flag", "flag", "memo"),
)
M_ROW = {key: 4 + i for i, (key, *_rest) in enumerate(MONTHLY_ROWS)}


def monthly_formula(key: str, m: int) -> object:
    """Excel formula for row ``key`` in model month ``m`` (0 = pre-opening)."""
    R = M_ROW
    c, p = mcol(m), mcol(m - 1) if m else None
    first = mcol(0)
    mo = f"{c}${R['month']}"

    def r(k):
        return f"{c}{R[k]}"

    f = {
        "month": 0 if m == 0 else f"={p}{R['month']}+1",
        "date": "=EDATE(Model_Start,-1)" if m == 0 else f"=EDATE({p}{R['date']},1)",
        "year": f"=IF({mo}=0,0,ROUNDUP({mo}/12,0))",
        "in_horizon": f"=IF({mo}<=Horizon_Months,1,0)",
        "days": f"=DAY(EOMONTH({c}${R['date']},0))",
        "season": f"=IF({mo}=0,0,INDEX(Seasonality,MONTH({c}${R['date']})))",
        "ramp": f"=IF({mo}=0,0,IF(Ramp_Months<=1,1,MIN(1,Ramp_Start+(1-Ramp_Start)*({mo}-1)/(Ramp_Months-1))))",
        "visitors_day": f"=Traffic_Plan*{r('ramp')}*{r('season')}",
        "visitors": f"={r('visitors_day')}*{r('days')}",
        "receipts": f"={r('visitors')}*Conversion",
        "ticket": f"=IF({mo}=0,0,Avg_Ticket*(1+Ticket_Growth)^({r('year')}-1))",
        "index": f"=IF({mo}=0,0,(1+Opex_Indexation)^({r('year')}-1))",
        "revenue": f"={r('receipts')}*{r('ticket')}",
        "cogs": f"=-{r('revenue')}*(1-Gross_Margin)",
        "gross_profit": f"={r('revenue')}+{r('cogs')}",
        "shrinkage": f"=-{r('revenue')}*Shrinkage",
        "gp_net": f"={r('gross_profit')}+{r('shrinkage')}",
        "rent": f"=-Rent_Month*{r('index')}",
        "payroll": f"=-Payroll_Month*{r('index')}",
        "utilities": f"=-Utilities_Month*{r('index')}",
        "marketing": f"=-(Mkt_Monthly*{r('index')}+IF({mo}=1,Mkt_Launch,0))",
        "acquiring": f"=-{r('revenue')}*Acquiring_Rate",
        "other": f"=-Other_Opex*{r('index')}",
        "opex": f"=SUM({r('rent')}:{r('other')})",
        "ebitda": f"={r('gp_net')}+{r('opex')}",
        "ebitda_margin": f"=IF({r('revenue')}=0,0,{r('ebitda')}/{r('revenue')})",
        "da": f"=IF(AND({mo}>=1,{mo}<=Depr_Months),-(Capex_Fitout+Capex_Equipment)/Depr_Months,0)",
        "ebt": f"={r('ebitda')}+{r('da')}",
        "cum_ebt": f"={r('ebt')}" if m == 0 else f"={p}{R['cum_ebt']}+{r('ebt')}",
        # Tax due to date = rate x highest cumulative profit so far (losses carried forward);
        # this month's tax is the increase in tax due.
        "tax": (f"=-Tax_Rate*MAX(0,{r('cum_ebt')})" if m == 0 else
                f"=-Tax_Rate*(MAX(0,MAX(${first}{R['cum_ebt']}:{r('cum_ebt')}))"
                f"-MAX(0,MAX(${first}{R['cum_ebt']}:{p}{R['cum_ebt']})))"),
        "net_profit": f"={r('ebt')}+{r('tax')}",
        "cf_ebitda": f"={r('ebitda')}",
        "cf_tax": f"={r('tax')}",
        "capex_fe": f"=IF({mo}=0,-(Capex_Fitout+Capex_Equipment),0)",
        "capex_inv": f"=IF({mo}=0,-Capex_Inventory,0)",
        "deposit": f"=IF({mo}=0,-Capex_Deposit,0)",
        "fcf_op": f"=SUM({r('cf_ebitda')}:{r('deposit')})*{r('in_horizon')}",
        "rec_inv": f'=IF(AND(Terminal_Recovery="Yes",{mo}=Horizon_Months),Capex_Inventory,0)',
        "rec_dep": f'=IF(AND(Terminal_Recovery="Yes",{mo}=Horizon_Months),Capex_Deposit,0)',
        "fcf": f"={r('fcf_op')}+({r('rec_inv')}+{r('rec_dep')})*{r('in_horizon')}",
        "cum_fcf": f"={r('fcf')}" if m == 0 else f"={p}{R['cum_fcf']}+{r('fcf')}",
        "df": f"=1/(1+Discount_Rate_M)^{mo}",
        "dcf": f"={r('fcf')}*{r('df')}",
        "cum_dcf": f"={r('dcf')}" if m == 0 else f"={p}{R['cum_dcf']}+{r('dcf')}",
        "cum_op": f"={r('fcf_op')}" if m == 0 else f"={p}{R['cum_op']}+{r('fcf_op')}",
        "cum_dcf_op": (f"={r('fcf_op')}*{r('df')}" if m == 0 else
                       f"={p}{R['cum_dcf_op']}+{r('fcf_op')}*{r('df')}"),
        "paid": f"=IF(AND({mo}>=1,{r('cum_op')}>=0),1,0)",
        "paid_disc": f"=IF(AND({mo}>=1,{r('cum_dcf_op')}>=0),1,0)",
    }
    return f[key]


def build_monthly(wb) -> None:
    """36-month projection; month 0 is the pre-opening investment month."""
    ws = wb.create_sheet("Monthly")
    sheet_title(ws, "Monthly projection")
    put(ws, "B2", '="Scenario: "&Active_Scenario&"   |   USD, excl. VAT. Costs are negative."',
        italic=True, color=GREY)
    last = M_FIRST_COL + N_MONTHS
    ws.column_dimensions["B"].width = 44
    ws.column_dimensions["C"].width = 10
    for col in range(M_FIRST_COL, last + 1):
        ws.column_dimensions[get_column_letter(col)].width = 10.5

    for key, label, unit, fmt, style in MONTHLY_ROWS:
        row = M_ROW[key]
        if style == "section":
            section(ws, row, label, 2, last)
            continue
        bold = style in ("total", "key", "head")
        italic = style == "memo"
        color = GREY if italic else BLACK
        put(ws, f"B{row}", label, bold=bold, italic=italic, color=color)
        put(ws, f"C{row}", unit, italic=True, color=GREY, align="center")
        for m in range(N_MONTHS + 1):
            ref = f"{mcol(m)}{row}"
            put(ws, ref, monthly_formula(key, m), fmt, bold=bold, italic=italic,
                color=BLUE if (key == "month" and m == 0) else (GREY if italic else None),
                align="center" if style == "head" else None)
            cell = ws[ref]
            if style == "head":
                cell.fill = HEADER_FILL
            elif m == 0:
                cell.fill = PRE_FILL
            if style in ("total", "key"):
                cell.border = TOP_BORDER
        if style == "head":
            for col in "BC":
                ws[f"{col}{row}"].fill = HEADER_FILL

    put(ws, f"{mcol(0)}3", "Pre-opening", italic=True, color=GREY, align="center")
    ws.freeze_panes = f"{mcol(0)}{M_ROW['date'] + 1}"
    ws.sheet_properties.tabColor = NAVY
    page_setup(ws)


def mrange(key: str, first: int = 0, absolute: bool = True) -> str:
    """Monthly!$D$r:$AN$r style range for a Monthly row (months ``first``..36)."""
    d = "$" if absolute else ""
    r = M_ROW[key]
    return f"Monthly!{d}{mcol(first)}{d}{r}:{d}{mcol(N_MONTHS)}{d}{r}"


# --------------------------------------------------------------------------- #
# Engine (scenario and sensitivity cases)
# --------------------------------------------------------------------------- #
def engine_cases() -> list[dict]:
    """Every case the Engine re-calculates, with its seven driver multipliers.

    Multipliers scale the active-scenario Monthly lines: traffic, conversion and
    ticket scale revenue; gross margin scales the margin %; rent scales rent and
    the deposit; payroll scales payroll; fit-out & equipment scales that capex
    and its depreciation.
    """
    cases = [{"id": "REF", "label": "Reference: active scenario", "mult": [1] * 7}]

    for i, name in enumerate(A.SCENARIOS):
        col = SCEN_COLS[i]
        r = SCEN_FIRST_DRIVER_ROW
        cases.append({
            "id": f"SCN_{name[:4].upper()}", "label": f"Scenario: {name}",
            "mult": [f"=Scenarios!${col}${r}/Traffic_Plan", f"=Scenarios!${col}${r + 1}/Conversion",
                     f"=Scenarios!${col}${r + 2}/Avg_Ticket", f"=Scenarios!${col}${r + 3}/Gross_Margin",
                     f"=Scenarios!${col}${r + 4}/Rent_Rate", 1, 1],
        })

    for d, (label, idx) in enumerate(TORNADO):
        for k, col in enumerate(SENS_SHOCK_COLS):
            mult = [1] * 7
            mult[idx] = f"=1+Sensitivity!${col}${SENS_HEADER_ROW}"
            cases.append({"id": f"SENS_{d + 1}_{k + 1}", "label": f"Sensitivity: {label}", "mult": mult,
                          "sens": (d, col)})

    for i in range(len(A.GRID_TRAFFIC)):
        for j in range(len(A.GRID_TICKET)):
            tcol = get_column_letter(3 + j)
            mult = [1] * 7
            mult[0] = f"=Sensitivity!$B${GRID_HEADER_ROW + 1 + i}/Traffic_Plan"
            mult[2] = f"=Sensitivity!${tcol}${GRID_HEADER_ROW}/Avg_Ticket"
            cases.append({"id": f"GRID_{i + 1}_{j + 1}", "label": "Traffic x ticket grid", "mult": mult,
                          "grid": (i, j)})

    for n, case in enumerate(cases):
        case["row"] = ENGINE_FIRST_ROW + n * ENGINE_BLOCK
    return cases


def engine_formula(line: int, m: int, b: int) -> str:
    """Formula for Engine line ``line`` (0-7) of the case block starting at row ``b``."""
    R = M_ROW
    c, p, mc = ecol(m), ecol(m - 1) if m else None, mcol(m)
    first = ecol(0)

    def mon(k):
        return f"Monthly!{mc}${R[k]}"

    def own(offset, col=c):
        return f"{col}{b + offset}"

    if line == 0:
        return f"={mon('revenue')}*$E${b}*$F${b}*$G${b}"
    if line == 1:
        return (f"={own(0)}*(Gross_Margin*$H${b}-Shrinkage-Acquiring_Rate)"
                f"+{mon('rent')}*$I${b}+{mon('payroll')}*$J${b}"
                f"+{mon('utilities')}+{mon('marketing')}+{mon('other')}")
    if line == 2:
        return f"={own(1)}+{mon('da')}*$K${b}"
    if line == 3:
        return f"=SUM(${first}{b + 2}:{own(2)})"
    if line == 4:
        if m == 0:
            return f"=-Tax_Rate*MAX(0,{own(3)})"
        return (f"=-Tax_Rate*(MAX(0,MAX(${first}{b + 3}:{own(3)}))"
                f"-MAX(0,MAX(${first}{b + 3}:{own(3, p)})))")
    if line == 5:
        return (f"=({own(1)}+{own(4)}+{mon('capex_fe')}*$K${b}+{mon('capex_inv')}"
                f"+{mon('deposit')}*$I${b}+{mon('rec_inv')}+{mon('rec_dep')}*$I${b})"
                f"*{mon('in_horizon')}")
    if line == 6:  # payback basis: cumulative FCF without the horizon-end recovery
        recovery = f"({mon('rec_inv')}+{mon('rec_dep')}*$I${b})*{mon('in_horizon')}"
        return f"={own(5)}-{recovery}" if m == 0 else f"={own(6, p)}+{own(5)}-{recovery}"
    return f"=IF(AND(Monthly!{mc}${R['month']}>=1,{own(6)}>=0),1,0)"


def payback_formula(flags: str, cum_m1: str, cum_m0: str, fallback: str = "n/a") -> str:
    """Interpolated payback month.

    ``flags``: paid-back flags for months 1..36; ``cum_m1``: cumulative cash flow
    for months 1..36; ``cum_m0``: the same row from month 0, so that INDEX(cum_m0, p)
    is the balance one month *before* month p. Payback = p - C(p) / (C(p) - C(p-1)).
    """
    p = f"MATCH(1,{flags},0)"
    return (f'=IFERROR({p}-INDEX({cum_m1},{p})/(INDEX({cum_m1},{p})-INDEX({cum_m0},{p})),'
            f'"{fallback}")')


def engine_outputs(b: int) -> dict:
    """Result formulas written on the first row of each case block."""
    rng = {i: f"${ecol(0)}{b + i}:${ecol(N_MONTHS)}{b + i}" for i in range(8)}
    m1 = {i: f"${ecol(1)}{b + i}:${ecol(N_MONTHS)}{b + i}" for i in range(8)}
    years = mrange("year")
    out = {
        "NPV": f"=SUMPRODUCT({rng[5]},{mrange('df')})",
        "IRR (annual)": f'=IFERROR((1+IRR({rng[5]},0.02))^12-1,"n/a")',
        "Payback (months)": payback_formula(m1[7], m1[6], rng[6]),
        "Peak funding": f"=-MIN({rng[6]})",
    }
    for y in (1, 2, 3):
        out[f"Revenue Y{y}"] = f"=SUMIF({years},{y},{rng[0]})"
        out[f"EBITDA Y{y}"] = f"=SUMIF({years},{y},{rng[1]})"
    return out


def build_engine(wb, cases: list[dict]) -> None:
    ws = wb.create_sheet("Engine")
    sheet_title(ws, "Engine - calculation block (no inputs here)",
                "Re-runs the monthly cash flow for each scenario and sensitivity case. "
                "Multipliers scale the active-scenario lines on the Monthly sheet.")
    header_row(ws, 4, {"B": "Case", "C": "Description", "D": "Line"})
    for name, col in zip(MULTIPLIERS, MULT_COLS):
        header_row(ws, 4, {col: name})
    for m in range(N_MONTHS + 1):
        put(ws, f"{ecol(m)}4", f"=Monthly!{mcol(m)}${M_ROW['month']}", "month", bold=True,
            fill=HEADER_FILL, align="center")
        put(ws, f"{ecol(m)}5", f"=Monthly!{mcol(m)}${M_ROW['date']}", "date", italic=True, align="center")
    for name, col in OUT_COL.items():
        header_row(ws, 4, {col: name})

    widths = {"B": 13, "C": 26, "D": 18}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    for col in MULT_COLS:
        ws.column_dimensions[col].width = 9
    for m in range(N_MONTHS + 1):
        ws.column_dimensions[ecol(m)].width = 10
    for col in OUT_COL.values():
        ws.column_dimensions[col].width = 12

    for case in cases:
        b = case["row"]
        put(ws, f"B{b}", case["id"], bold=True)
        put(ws, f"C{b}", case["label"])
        for value, col in zip(case["mult"], MULT_COLS):
            put(ws, f"{col}{b}", value, "factor", fill=PRE_FILL)
        for line, label in enumerate(ENGINE_LINES):
            put(ws, f"D{b + line}", label, italic=line in (3, 7), color=GREY if line in (3, 7) else BLACK)
            fmt = "flag" if line == 7 else "money"
            for m in range(N_MONTHS + 1):
                put(ws, f"{ecol(m)}{b + line}", engine_formula(line, m, b), fmt)
        for name, formula in engine_outputs(b).items():
            fmt = {"IRR (annual)": "pct", "Payback (months)": "months"}.get(name, "money")
            put(ws, f"{OUT_COL[name]}{b}", formula, fmt, bold=True, fill=HEADER_FILL)

    ws.freeze_panes = f"{ecol(0)}6"
    ws.sheet_properties.tabColor = "A6A6A6"
    page_setup(ws)


def engine_ref(case: dict, output: str) -> str:
    return f"Engine!${OUT_COL[output]}${case['row']}"


# --------------------------------------------------------------------------- #
# Scenarios
# --------------------------------------------------------------------------- #
def build_scenarios(wb, cases: list[dict]) -> None:
    ws = wb.create_sheet("Scenarios")
    sheet_title(ws, "Scenarios")
    put(ws, "B2", '="Active scenario: "&Active_Scenario&"   (switch it on the Inputs sheet)"',
        italic=True, color=GREY)
    for col, w in zip("BCDEFGH", (40, 14, 14, 14, 14, 16, 48)):
        ws.column_dimensions[col].width = w

    r = SCEN_HEADER_ROW
    header_row(ws, r, {"B": "Driver", "F": "Active", "G": "Unit", "H": "Rationale"})
    for name, col in zip(A.SCENARIOS, SCEN_COLS):
        header_row(ws, r, {col: name})
    for k, drv in enumerate(A.DRIVERS):
        row = SCEN_FIRST_DRIVER_ROW + k
        put(ws, f"B{row}", drv.label)
        for value, col in zip(drv.values, SCEN_COLS):
            put(ws, f"{col}{row}", value, drv.fmt, color=BLUE, fill=INPUT_FILL, border=BOX)
        put(ws, f"F{row}", f"=INDEX($C{row}:$E{row},Scenario_Index)", drv.fmt, bold=True, border=BOX)
        put(ws, f"G{row}", drv.unit, color=GREY)
        put(ws, f"H{row}", drv.note, color=GREY)
    # Highlight the active scenario column.
    ws.conditional_formatting.add(
        f"C{SCEN_HEADER_ROW}:E{SCEN_HEADER_ROW + len(A.DRIVERS)}",
        FormulaRule(formula=[f"C${SCEN_HEADER_ROW}=Active_Scenario"], font=Font(name=FONT, bold=True),
                    border=Border(left=Side(style="medium", color=NAVY), right=Side(style="medium", color=NAVY))))

    scen_cases = [c for c in cases if c["id"].startswith("SCN_")]
    r = 12
    section(ws, r, "Scenario comparison (live - recalculates with every input)", 2, 6)
    r += 1
    header_row(ws, r, {"B": "Metric", "F": "Unit"})
    for name, col in zip(A.SCENARIOS, SCEN_COLS):
        header_row(ws, r, {col: name})
    d = SCEN_FIRST_DRIVER_ROW  # driver rows: traffic, conversion, ticket, margin, rent

    def be_traffic(col):
        fixed = f"(Payroll_Month+Utilities_Month+Mkt_Monthly+Other_Opex+Store_Area*{col}${d + 4})"
        return f"={fixed}/({col}${d + 3}-Shrinkage-Acquiring_Rate)/(365/12*{col}${d + 1}*{col}${d + 2})"

    def payback(k):
        ref = engine_ref(k, "Payback (months)")
        return f'=IF(ISNUMBER({ref}),{ref},"> "&Horizon_Months)'

    # Each lambda gets (scenario column, engine case, own row).
    rows = (
        ("Revenue, Year 1", "USD", "money", lambda c, k, r: f"={engine_ref(k, 'Revenue Y1')}"),
        ("Revenue, Year 3", "USD", "money", lambda c, k, r: f"={engine_ref(k, 'Revenue Y3')}"),
        ("EBITDA, Year 1", "USD", "money", lambda c, k, r: f"={engine_ref(k, 'EBITDA Y1')}"),
        ("EBITDA, Year 3", "USD", "money", lambda c, k, r: f"={engine_ref(k, 'EBITDA Y3')}"),
        ("EBITDA margin, Year 3", "%", "pct", lambda c, k, r: f"=IF({c}{r - 3}=0,0,{c}{r - 1}/{c}{r - 3})"),
        ("Break-even traffic", "visitors/day", "int", lambda c, k, r: be_traffic(c)),
        ("Margin of safety vs planned traffic", "%", "pct", lambda c, k, r: f"=1-{c}{r - 1}/{c}${d}"),
        ("Total investment", "USD", "money",
         lambda c, k, r: f"=Capex_Fitout+Capex_Equipment+Capex_Inventory+Store_Area*{c}${d + 4}*Deposit_Months"),
        ("Peak funding requirement", "USD", "money", lambda c, k, r: f"={engine_ref(k, 'Peak funding')}"),
        ("NPV", "USD", "money", lambda c, k, r: f"={engine_ref(k, 'NPV')}"),
        ("IRR (annual)", "%", "pct", lambda c, k, r: f"={engine_ref(k, 'IRR (annual)')}"),
        ("Payback (months from opening)", "months", "months", lambda c, k, r: payback(k)),
    )
    for label, unit, fmt, fn in rows:
        r += 1
        put(ws, f"B{r}", label, bold=label in ("NPV", "Payback (months from opening)"))
        put(ws, f"F{r}", unit, color=GREY)
        for col, case in zip(SCEN_COLS, scen_cases):
            put(ws, f"{col}{r}", fn(col, case, r), fmt, border=BOX, align="right")
    put(ws, f"B{r + 2}", "Break-even and margin of safety: mature store (after ramp-up), Year-1 prices and costs, "
                         "average month (seasonality = 1.00), EBITDA = 0.", italic=True, color=GREY)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "FFC000"
    page_setup(ws)


# --------------------------------------------------------------------------- #
# Summary
# --------------------------------------------------------------------------- #
def build_summary(wb) -> None:
    ws = wb.create_sheet("Summary")
    sheet_title(ws, "Summary - store opening business case")
    put(ws, "B2", '="Scenario: "&Active_Scenario&"   |   Model checks: "&Checks_Status&"   |   USD, excl. VAT"',
        italic=True, color=GREY)
    for col, w in zip("BCDEFG", (40, 14, 20, 14, 14, 60)):
        ws.column_dimensions[col].width = w

    section(ws, 4, "Profit & loss by model year", 2, 6)
    header_row(ws, 5, {"B": "USD"})
    for col, y in zip("CDE", (1, 2, 3)):
        put(ws, f"{col}5", y, "year", bold=True, fill=HEADER_FILL, align="center", color=BLACK)
    header_row(ws, 5, {"F": "Total"})
    years = mrange("year")
    pl_rows = (("Revenue", "revenue"), ("Gross profit after shrinkage", "gp_net"), ("Total OPEX", "opex"),
               ("EBITDA", "ebitda"), ("EBITDA margin", None), ("Net profit", "net_profit"),
               ("Free cash flow", "fcf"))
    row = 6
    for label, key in pl_rows:
        bold = label in ("EBITDA", "Net profit", "Revenue")
        put(ws, f"B{row}", label, bold=bold)
        for col in "CDEF":
            if key is None:
                f = f"=IF({col}6=0,0,{col}9/{col}6)"
                fmt = "pct"
            elif col == "F":
                f = f"=SUM(C{row}:E{row})"
                fmt = "money"
            else:
                f = f"=SUMIF({years},{col}$5,{mrange(key)})"
                fmt = "money"
            put(ws, f"{col}{row}", f, fmt, bold=bold, border=BOX)
        row += 1

    payback = payback_formula(mrange("paid", 1), mrange("cum_op", 1), mrange("cum_op"), "Not reached")
    payback_d = payback_formula(mrange("paid_disc", 1), mrange("cum_dcf_op", 1), mrange("cum_dcf_op"), "Not reached")

    blocks = (
        ("Investment & returns", (
            ("Capex_Total_S", "Total investment", "=Capex_Total", "money", "USD",
             "Fit-out, equipment, initial inventory and rent deposit (month 0)."),
            ("Peak_Funding", "Peak funding requirement", f"=-MIN({mrange('cum_op')})", "money", "USD",
             "Deepest point of cumulative cash flow = cash the owner must have available."),
            ("Payback_Months", "Payback period", payback, "months", "months from opening",
             "First month when cumulative cash flow (before horizon-end recovery) turns positive, interpolated."),
            ("Payback_Disc", "Discounted payback", payback_d, "months", "months from opening",
             "Same, using cash flows discounted at the discount rate."),
            ("NPV_Value", "NPV", f"=SUM({mrange('dcf')})", "money", "USD",
             "Sum of monthly cash flows discounted at the monthly equivalent of the annual rate."),
            ("IRR_Annual", "IRR (annualised)", f'=IFERROR((1+IRR({mrange("fcf")},0.02))^12-1,"n/a")', "pct", "per year",
             "Monthly IRR of the cash-flow series, compounded to a year."),
        )),
        ("Break-even (mature store)", (
            ("CM_Pct", "Contribution margin", "=Gross_Margin-Shrinkage-Acquiring_Rate", "pct", "% of revenue",
             "Gross margin less revenue-linked costs (shrinkage, acquiring)."),
            ("Fixed_Month", "Fixed costs per month", "=Rent_Month+Payroll_Month+Utilities_Month+Mkt_Monthly+Other_Opex",
             "money", "USD/month", "Year-1 level, excluding the one-off launch campaign."),
            ("BE_Revenue", "Break-even revenue", "=Fixed_Month/CM_Pct", "money", "USD/month",
             "Revenue at which EBITDA = 0."),
            ("BE_Traffic", "Break-even traffic", "=BE_Revenue/(365/12*Conversion*Avg_Ticket)", "int", "visitors/day",
             "Break-even revenue / (days x conversion x average ticket)."),
            ("Plan_Traffic_S", "Planned traffic", "=Traffic_Plan", "int", "visitors/day", "Active scenario."),
            ("Safety_Margin", "Margin of safety", "=1-BE_Traffic/Traffic_Plan", "pct", "% of planned traffic",
             "How far traffic can fall below plan before the store stops covering its costs."),
        )),
    )
    row += 1
    for title, items in blocks:
        section(ws, row, title, 2, 6)
        row += 1
        header_row(ws, row, {"B": "Metric", "C": "Value", "D": "Unit"})
        ws.merge_cells(f"E{row}:F{row}")
        header_row(ws, row, {"E": "How it is calculated"})
        row += 1
        for name, label, formula, fmt, unit, note in items:
            key = name in ("Payback_Months", "NPV_Value", "IRR_Annual", "BE_Traffic")
            put(ws, f"B{row}", label, bold=key)
            put(ws, f"C{row}", formula, fmt, bold=key, border=BOX, align="right")
            put(ws, f"D{row}", unit, color=GREY)
            put(ws, f"E{row}", note, color=GREY)
            add_name(wb, name, f"Summary!$C${row}")
            row += 1
        row += 1

    chart = LineChart()
    chart.title = "Cumulative free cash flow before horizon-end recovery, USD"
    chart.height, chart.width = 7.5, 17
    chart.legend = None
    data = Reference(wb["Monthly"], min_col=M_FIRST_COL, max_col=M_FIRST_COL + N_MONTHS, min_row=M_ROW["cum_op"])
    chart.add_data(data, from_rows=True, titles_from_data=False)
    chart.set_categories(Reference(wb["Monthly"], min_col=M_FIRST_COL, max_col=M_FIRST_COL + N_MONTHS,
                                   min_row=M_ROW["date"]))
    chart.series[0].graphicalProperties.line.solidFill = "2A78D6"
    chart.series[0].graphicalProperties.line.width = 22000
    chart.series[0].smooth = False
    chart.y_axis.numFmt = '#,##0'
    chart.x_axis.tickLblPos = "low"
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    ws.add_chart(chart, f"B{row + 1}")

    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "70AD47"
    page_setup(ws, landscape=False)


# --------------------------------------------------------------------------- #
# Sensitivity
# --------------------------------------------------------------------------- #
def build_sensitivity(wb, cases: list[dict]) -> None:
    ws = wb.create_sheet("Sensitivity")
    sheet_title(ws, "Sensitivity analysis",
                "One driver changes at a time; everything else stays at the active scenario. "
                "Yellow cells (shock sizes, grid axes) can be edited.")
    for col, w in zip("BCDEFGHIJKL", (26, 13, 13, 13, 13, 13, 14, 14, 13, 8, 12)):
        ws.column_dimensions[col].width = w

    ref_case = cases[0]
    sens = {c["sens"]: c for c in cases if "sens" in c}
    grid = {c["grid"]: c for c in cases if "grid" in c}

    h = SENS_HEADER_ROW
    section(ws, h - 1, "NPV by driver change, USD", 2, 11)
    header_row(ws, h, {"B": "Driver", "E": "Base", "H": "NPV change at low shock", "I": "NPV change at high shock",
                       "J": "Swing (high - low)", "K": "Rank"})
    for col, shock in zip(SENS_SHOCK_COLS, A.SENSITIVITY_SHOCKS):
        put(ws, f"{col}{h}", shock, "shock", bold=True, fill=INPUT_FILL, color=BLUE, align="center", border=BOX)
    put(ws, f"L{h}", "Sort key", italic=True, color=GREY, align="center")
    first, last = SENS_FIRST_ROW, SENS_FIRST_ROW + len(TORNADO) - 1
    for d, (label, _) in enumerate(TORNADO):
        row = first + d
        put(ws, f"B{row}", label)
        for col in SENS_SHOCK_COLS:
            put(ws, f"{col}{row}", f"={engine_ref(sens[(d, col)], 'NPV')}", "money", border=BOX)
        put(ws, f"E{row}", f"={engine_ref(ref_case, 'NPV')}", "money", border=BOX, bold=True)
        put(ws, f"H{row}", f"=C{row}-E{row}", "money", border=BOX)
        put(ws, f"I{row}", f"=G{row}-E{row}", "money", border=BOX)
        put(ws, f"J{row}", f"=ABS(G{row}-C{row})", "money", border=BOX)
        put(ws, f"K{row}", f"=RANK(L{row},$L${first}:$L${last})", "int", border=BOX, align="center")
        # Tiny row-based offset breaks ties so the sorted table below never repeats a driver.
        put(ws, f"L{row}", f"=J{row}-ROW()/1000000", "money", italic=True, color=GREY)
    put(ws, f"B{last + 1}", "Traffic, conversion and ticket multiply into revenue, so equal % changes give equal NPV "
                            "changes. Costs (rent, payroll, capex) hurt NPV when they go up.",
        italic=True, color=GREY)

    # Sorted table feeding the tornado chart.
    t = last + 3
    section(ws, t, "Tornado (sorted by swing)", 2, 5)
    header_row(ws, t + 1, {"B": "Driver", "E": "#"})
    put(ws, f"C{t + 1}", f'="Low ("&TEXT(C{h},"+0%;-0%")&")"', bold=True, fill=HEADER_FILL, align="center")
    put(ws, f"D{t + 1}", f'="High ("&TEXT(G{h},"+0%;-0%")&")"', bold=True, fill=HEADER_FILL, align="center")
    for k in range(len(TORNADO)):
        row = t + 2 + k
        put(ws, f"E{row}", f"=MATCH(LARGE($L${first}:$L${last},{k + 1}),$L${first}:$L${last},0)", "int",
            color=GREY, align="center")
        put(ws, f"B{row}", f"=INDEX($B${first}:$B${last},E{row})")
        put(ws, f"C{row}", f"=INDEX($H${first}:$H${last},E{row})", "money", border=BOX)
        put(ws, f"D{row}", f"=INDEX($I${first}:$I${last},E{row})", "money", border=BOX)
    t_last = t + 1 + len(TORNADO)

    chart = BarChart()
    chart.type = "bar"
    chart.grouping = "clustered"
    chart.overlap = 100
    chart.gapWidth = 40
    chart.title = "NPV change vs base, USD"
    chart.height, chart.width = 8, 16
    chart.add_data(Reference(ws, min_col=3, max_col=4, min_row=t + 1, max_row=t_last), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=2, min_row=t + 2, max_row=t_last))
    for series, color in zip(chart.series, ("EB6834", "2A78D6")):
        series.graphicalProperties.solidFill = color
        series.graphicalProperties.line.noFill = True
        series.invertIfNegative = False  # Excel paints negative bars white otherwise
    chart.x_axis.scaling.orientation = "maxMin"  # biggest swing on top
    chart.y_axis.crosses = "max"                  # keep the value axis at the bottom
    chart.x_axis.tickLblPos = "low"
    chart.y_axis.numFmt = '#,##0'
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.legend.position = "r"
    ws.add_chart(chart, "H15")

    # Two-way grid: traffic x ticket -> payback.
    g = GRID_HEADER_ROW
    section(ws, g - 1, "Payback, months from opening: planned traffic (rows) x average ticket, USD (columns)", 2, 7)
    put(ws, f"B{g}", "Traffic, visitors/day", bold=True, fill=HEADER_FILL, align="center")
    for j, ticket in enumerate(A.GRID_TICKET):
        put(ws, f"{get_column_letter(3 + j)}{g}", ticket, "usd0", bold=True, fill=INPUT_FILL, color=BLUE,
            align="center", border=BOX)
    for i, traffic in enumerate(A.GRID_TRAFFIC):
        row = g + 1 + i
        put(ws, f"B{row}", traffic, "int", bold=True, fill=INPUT_FILL, color=BLUE, align="center", border=BOX)
        for j in range(len(A.GRID_TICKET)):
            ref = engine_ref(grid[(i, j)], "Payback (months)")
            put(ws, f"{get_column_letter(3 + j)}{row}", f'=IF(ISNUMBER({ref}),{ref},"> "&Horizon_Months)',
                "months", align="center", border=BOX)
    g_last = g + len(A.GRID_TRAFFIC)
    grid_ref = f"C{g + 1}:{get_column_letter(2 + len(A.GRID_TICKET))}{g_last}"
    ws.conditional_formatting.add(grid_ref, FormulaRule(formula=[f"ISTEXT(C{g + 1})"], fill=ERR_FILL,
                                                        stopIfTrue=True))
    ws.conditional_formatting.add(grid_ref, ColorScaleRule(
        start_type="num", start_value=12, start_color="C6EFCE",
        mid_type="num", mid_value=24, mid_color="FFEB9C",
        end_type="num", end_value=36, end_color="FFC7CE"))
    ws.conditional_formatting.add(grid_ref, FormulaRule(
        formula=[f"AND($B{g + 1}=Traffic_Plan,C${g}=Avg_Ticket)"],
        font=Font(name=FONT, bold=True), border=Border(left=Side(style="medium"), right=Side(style="medium"),
                                                       top=Side(style="medium"), bottom=Side(style="medium"))))
    put(ws, f"B{g_last + 1}", "Bold box = active scenario. Green = fast payback, red = slow; "
                              "'> 36' = not paid back within the horizon.", italic=True, color=GREY)

    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "70AD47"
    page_setup(ws)


# --------------------------------------------------------------------------- #
# Checks
# --------------------------------------------------------------------------- #
def build_checks(wb, cases: list[dict]) -> None:
    ws = wb.create_sheet("Checks")
    sheet_title(ws, "Model integrity checks")
    for col, w in zip("BCDEF", (5, 70, 16, 12, 12)):
        ws.column_dimensions[col].width = w

    put(ws, "C3", "Overall status", bold=True)
    put(ws, "F3", '=IF(COUNTIF(F6:F40,"ERROR")=0,"OK","ERROR")', bold=True, align="center")
    status_format(ws, "F3")
    add_name(wb, "Checks_Status", "Checks!$F$3")
    header_row(ws, 5, {"B": "#", "C": "Check", "D": "Difference", "E": "Tolerance", "F": "Status"})

    ref_case = cases[0]
    scen_npv = [engine_ref(c, "NPV") for c in cases if c["id"].startswith("SCN_")]
    s = mrange  # Monthly row ranges, all 37 columns
    last_engine = cases[-1]["row"] + ENGINE_BLOCK
    rates = ("Conversion", "Gross_Margin", "Shrinkage", "Tax_Rate", "Card_Share", "Ramp_Start",
             "Acquiring_Fee", "Payroll_Tax")
    in_range = ",".join(f"{n}>=0,{n}<=1" for n in rates)
    def sc(col, k):
        return f"Scenarios!${col}${SCEN_FIRST_DRIVER_ROW + k}"

    # Revenue drivers and margin rise from Pessimistic to Optimistic; rent (k = 4) falls.
    ordered = [f"{sc('C', k)}<={sc('D', k)},{sc('D', k)}<={sc('E', k)}" for k in range(4)]
    ordered.append(f"{sc('C', 4)}>={sc('D', 4)},{sc('D', 4)}>={sc('E', 4)}")

    checks = (
        ("Active scenario name is recognised", "=IF(ISNUMBER(Scenario_Index),0,1)", 0),
        ("Seasonality indices average 1.00", "=AVERAGE(Seasonality)-1", 0.0005),
        ("All rates are between 0% and 100%", f"=IF(AND({in_range}),0,1)", 0),
        ("Horizon is a whole number of months between 12 and 36",
         "=IF(AND(Horizon_Months>=12,Horizon_Months<=36,INT(Horizon_Months)=Horizon_Months),0,1)", 0),
        ("Scenario ordering: Pessimistic <= Base <= Optimistic (rent reversed)",
         f"=IF(AND({','.join(ordered)}),0,1)", 0),
        ("Revenue = receipts x average ticket (36-month total)",
         f"=SUM({s('revenue')})-SUMPRODUCT({s('receipts')},{s('ticket')})", 0.01),
        ("P&L bridge: revenue + COGS + shrinkage + OPEX = EBITDA",
         f"=SUM({s('revenue')})+SUM({s('cogs')})+SUM({s('shrinkage')})+SUM({s('opex')})-SUM({s('ebitda')})", 0.01),
        ("Total OPEX = sum of OPEX lines",
         f"=SUM({s('opex')})-SUM({s('rent')},{s('payroll')},{s('utilities')},{s('marketing')},{s('acquiring')},{s('other')})",
         0.01),
        ("Net profit = EBITDA + depreciation + income tax",
         f"=SUM({s('net_profit')})-SUM({s('ebitda')},{s('da')},{s('tax')})", 0.01),
        ("Income tax is never a refund (no positive tax lines)", f'=COUNTIF({s("tax")},">0.001")', 0),
        ("Investment in cash flow = total investment on Inputs",
         f"=-SUM({s('capex_fe')},{s('capex_inv')},{s('deposit')})-Capex_Total", 0.01),
        ("Cumulative cash flow at month 36 = sum of monthly cash flows",
         f"=Monthly!${mcol(N_MONTHS)}${M_ROW['cum_fcf']}-SUM({s('fcf')})", 0.01),
        ("NPV (discount factors) = Excel NPV() function",
         f"=NPV_Value-(Monthly!${mcol(0)}${M_ROW['fcf']}+NPV(Discount_Rate_M,{mrange('fcf', first=1)}))", 0.01),
        ("Engine reference case NPV = Summary NPV", f"={engine_ref(ref_case, 'NPV')}-NPV_Value", 0.01),
        ("Engine reference case cumulative cash flow = Monthly",
         f"=Engine!${ecol(N_MONTHS)}${ref_case['row'] + 6}-Monthly!${mcol(N_MONTHS)}${M_ROW['cum_op']}", 0.01),
        ("Engine scenario case for the active scenario = Summary NPV",
         f"=CHOOSE(Scenario_Index,{','.join(scen_npv)})-NPV_Value", 0.01),
        ("Sensitivity base NPV = Summary NPV", f"=Sensitivity!$E${SENS_FIRST_ROW}-NPV_Value", 0.01),
        # Rows showing Checks_Status (Inputs row 4, Summary row 2) are skipped to avoid a circular reference.
        ("No error values (#REF!, #DIV/0!, #VALUE!...) on any sheet",
         f"=SUMPRODUCT(--ISERROR(Inputs!B5:E120))+SUMPRODUCT(--ISERROR(Scenarios!B1:H40))"
         f"+SUMPRODUCT(--ISERROR(Monthly!B1:{mcol(N_MONTHS)}{max(M_ROW.values())}))"
         f"+SUMPRODUCT(--ISERROR(Summary!B4:G40))+SUMPRODUCT(--ISERROR(Sensitivity!B1:L40))"
         f"+SUMPRODUCT(--ISERROR(Engine!B1:{OUT_COL['Peak funding']}{last_engine}))", 0),
    )
    for n, (label, formula, tol) in enumerate(checks, start=1):
        row = 5 + n
        put(ws, f"B{row}", n, "int", color=BLACK, align="center")
        put(ws, f"C{row}", label)
        put(ws, f"D{row}", formula, "#,##0.00;-#,##0.00;0.00", border=BOX)
        put(ws, f"E{row}", tol, "#,##0.0000", border=BOX, color=BLUE)
        put(ws, f"F{row}", f'=IF(ISERROR(D{row}),"ERROR",IF(ABS(D{row})<=E{row},"OK","ERROR"))',
            align="center", bold=True, border=BOX)
    status_format(ws, f"F6:F{5 + len(checks)}")
    put(ws, f"C{7 + len(checks)}", "Difference = calculated value minus expected value; OK when within tolerance.",
        italic=True, color=GREY)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "C00000"
    page_setup(ws, landscape=False)


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def build(out: Path = DEFAULT_OUT, scenario: str = "Base") -> Path:
    """Create the workbook and save it to ``out``. Returns the path."""
    if scenario not in A.SCENARIOS:
        raise ValueError(f"scenario must be one of {A.SCENARIOS}")
    assert [d.name for d in A.DRIVERS] == ["Traffic_Plan", "Conversion", "Avg_Ticket", "Gross_Margin", "Rent_Rate"]

    wb = Workbook()
    cases = engine_cases()
    build_inputs(wb, scenario)
    build_scenarios(wb, cases)
    build_monthly(wb)
    build_summary(wb)
    build_sensitivity(wb, cases)
    build_checks(wb, cases)
    build_engine(wb, cases)
    wb.calculation.fullCalcOnLoad = True

    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return out


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--out", type=Path, default=DEFAULT_OUT, help="output .xlsx path")
    parser.add_argument("--scenario", default="Base", choices=A.SCENARIOS, help="active scenario")
    args = parser.parse_args()
    print(f"Model written to {build(args.out, args.scenario)}")


if __name__ == "__main__":
    main()

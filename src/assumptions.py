"""Model assumptions for the illustrative home-goods store.

Every number here is fictional but plausible for a ~250 m2 home-goods store
in a city of ~1M people. All money values are in USD, excluding VAT.

This module is the single source of truth for the values written into the
yellow input cells of ``model/store_model.xlsx``. Edit a value here and run
``python run_all.py`` to rebuild every deliverable.
"""

from dataclasses import dataclass
from datetime import date


@dataclass(frozen=True)
class Input:
    """One labelled assumption on the Inputs sheet.

    Attributes:
        name: Excel defined name used by formulas (e.g. ``Tax_Rate``).
        label: Human-readable label shown in column B.
        value: Hard-coded input, or an Excel formula string starting with ``=``.
        unit: Unit shown next to the value.
        note: Rationale shown in the notes column.
        fmt: Number-format key (see ``build_model.FORMATS``).
        kind: ``input`` (yellow cell), ``link`` (pulled from Scenarios) or ``calc``.
        options: Allowed values for a drop-down list, if any.
    """

    name: str
    label: str
    value: object
    unit: str
    note: str
    fmt: str = "money"
    kind: str = "input"
    options: tuple = ()


SCENARIOS = ("Pessimistic", "Base", "Optimistic")


@dataclass(frozen=True)
class Driver:
    """A scenario driver: one row on the Scenarios sheet."""

    name: str
    label: str
    unit: str
    fmt: str
    values: tuple  # (Pessimistic, Base, Optimistic)
    note: str


DRIVERS = (
    Driver("Traffic_Plan", "Planned traffic at maturity", "visitors/day", "int",
           (430, 500, 560), "Neighbourhood mall / high-street unit of 250 m2."),
    Driver("Conversion", "Conversion (visitors to receipts)", "%", "pct",
           (0.23, 0.25, 0.26), "Home goods typically convert 20-30%."),
    Driver("Avg_Ticket", "Average ticket (Year 1)", "USD", "money2",
           (14.0, 15.0, 16.0), "Basket of 3-4 household items."),
    Driver("Gross_Margin", "Gross margin", "% of revenue", "pct",
           (0.36, 0.38, 0.39), "Category norm 35-45% incl. supplier bonuses."),
    Driver("Rent_Rate", "Rent rate", "USD/m2/month", "money2",
           (17.0, 16.0, 15.0), "Mid-market shopping centre."),
)

SEASONALITY = (  # calendar month -> traffic index; average must equal 1.00
    ("January", 0.85, "Post-holiday slowdown"),
    ("February", 0.85, ""),
    ("March", 0.95, ""),
    ("April", 1.00, "Spring cleaning / garden season"),
    ("May", 1.00, ""),
    ("June", 0.95, ""),
    ("July", 0.90, "Summer holidays"),
    ("August", 0.95, "Back-to-school"),
    ("September", 1.00, ""),
    ("October", 1.05, ""),
    ("November", 1.10, "Black Friday"),
    ("December", 1.40, "New Year gifts and decor"),
)

# (section title, inputs). Formula values may reference names defined above them.
SECTIONS = (
    ("General", (
        Input("Model_Start", "Opening month (first trading month)", date(2027, 3, 1),
              "date", "Investment is spent in the month before opening (month 0).", "[$-en-US]mmm yyyy"),
        Input("Horizon_Months", "Forecast horizon", 36, "months",
              "12-36. Cash flows after the horizon are excluded from NPV, IRR and payback.", "int"),
        Input("Store_Area", "Store area (total)", 250, "m2",
              "Of which ~200 m2 sales floor. Drives rent and utilities.", "int"),
        Input("Active_Scenario", "Active scenario", "Base", "list",
              "Switches the five scenario drivers (see Scenarios sheet).", "text",
              options=SCENARIOS),
        Input("Scenario_Index", "Active scenario position", "=MATCH(Active_Scenario,Scenarios!$C$4:$E$4,0)",
              "#", "Helper: column of the active scenario on the Scenarios sheet.", "int", "calc"),
    )),
    ("Traffic & sales", (
        Input("Traffic_Plan", "Planned traffic at maturity", "=Scenarios!$F$5", "visitors/day",
              "From Scenarios (active case).", "int", "link"),
        Input("Ramp_Start", "Traffic in opening month", 0.60, "% of plan",
              "New stores need time to build awareness.", "pct"),
        Input("Ramp_Months", "Ramp-up to planned traffic", 6, "months",
              "Linear ramp: month 1 = start %, month N = 100%.", "int"),
        Input("Conversion", "Conversion (visitors to receipts)", "=Scenarios!$F$6", "%",
              "From Scenarios (active case).", "pct", "link"),
        Input("Avg_Ticket", "Average ticket (Year 1)", "=Scenarios!$F$7", "USD",
              "From Scenarios (active case).", "money2", "link"),
        Input("Ticket_Growth", "Average ticket growth", 0.04, "per year",
              "USD inflation + mix. Applied in steps every 12 months.", "pct"),
    )),
    ("Margin", (
        Input("Gross_Margin", "Gross margin", "=Scenarios!$F$8", "% of revenue",
              "From Scenarios (active case).", "pct", "link"),
        Input("Shrinkage", "Shrinkage & write-offs", 0.015, "% of revenue",
              "Typical non-food level.", "pct"),
    )),
    ("Operating expenses (Year-1 level)", (
        Input("Rent_Rate", "Rent rate", "=Scenarios!$F$9", "USD/m2/month",
              "From Scenarios (active case).", "money2", "link"),
        Input("Rent_Month", "Rent per month", "=Store_Area*Rent_Rate", "USD/month",
              "Area x rate.", "money", "calc"),
        Input("Mgr_Count", "Store managers", 1, "FTE", "", "int"),
        Input("Mgr_Salary", "Manager gross salary", 900, "USD/month", "", "money"),
        Input("Staff_Count", "Sales associates / cashiers", 7, "FTE",
              "Two shifts, 7 days a week.", "int"),
        Input("Staff_Salary", "Associate gross salary", 550, "USD/month", "", "money"),
        Input("Payroll_Tax", "Employer payroll tax", 0.22, "% of salary",
              "Social contribution on gross payroll.", "pct"),
        Input("Payroll_Month", "Payroll per month",
              "=(Mgr_Count*Mgr_Salary+Staff_Count*Staff_Salary)*(1+Payroll_Tax)", "USD/month",
              "Headcount x salary x (1 + payroll tax).", "money", "calc"),
        Input("Utilities_Rate", "Utilities", 3.0, "USD/m2/month",
              "Electricity, heating, water.", "money2"),
        Input("Utilities_Month", "Utilities per month", "=Store_Area*Utilities_Rate", "USD/month",
              "", "money", "calc"),
        Input("Mkt_Launch", "Launch marketing", 6000, "USD, month 1",
              "Opening campaign: outdoor, social, local promo.", "money"),
        Input("Mkt_Monthly", "Ongoing marketing", 600, "USD/month",
              "About 1% of mature revenue.", "money"),
        Input("Acquiring_Fee", "Card acquiring fee", 0.013, "% of card sales", "", "pct"),
        Input("Card_Share", "Share of card payments", 0.80, "% of revenue", "", "pct"),
        Input("Acquiring_Rate", "Acquiring cost", "=Acquiring_Fee*Card_Share", "% of revenue",
              "Fee x card share.", "pct", "calc"),
        Input("Other_Opex", "Other OPEX", 900, "USD/month",
              "Security, cleaning, POS/IT, accounting, supplies.", "money"),
        Input("Opex_Indexation", "OPEX indexation", 0.04, "per year",
              "Applied to all fixed OPEX from Year 2, in yearly steps.", "pct"),
    )),
    ("Investment (month 0)", (
        Input("Capex_Fitout", "Fit-out / renovation", 30000, "USD", "About USD 120 per m2.", "money"),
        Input("Capex_Equipment", "Equipment", 25000, "USD",
              "Shelving, POS, anti-theft gates, CCTV, AC, signage.", "money"),
        Input("Capex_Inventory", "Initial inventory", 50000, "USD",
              "About 2 months of COGS (60-75 days of stock).", "money"),
        Input("Deposit_Months", "Rent deposit", 2, "months of rent", "", "int"),
        Input("Capex_Deposit", "Rent deposit amount", "=Rent_Month*Deposit_Months", "USD",
              "", "money", "calc"),
        Input("Capex_Total", "Total investment",
              "=Capex_Fitout+Capex_Equipment+Capex_Inventory+Capex_Deposit", "USD",
              "Cash needed before opening.", "money", "calc"),
        Input("Depr_Months", "Depreciation period", 60, "months",
              "Straight-line, fit-out and equipment. Used for income tax only.", "int"),
    )),
    ("Tax & valuation", (
        Input("Tax_Rate", "Corporate income tax", 0.18, "% of profit before tax",
              "Tax losses are carried forward.", "pct"),
        Input("Discount_Rate", "Discount rate", 0.20, "per year, USD",
              "Small retail business, emerging-market risk premium.", "pct"),
        Input("Discount_Rate_M", "Discount rate, monthly", "=(1+Discount_Rate)^(1/12)-1", "per month",
              "Annual rate converted to an equivalent monthly rate.", "pct2", "calc"),
        Input("Terminal_Recovery", "Recover inventory & deposit at horizon end", "Yes", "Yes/No",
              "Cash that physically returns. No going-concern value is added (conservative).",
              "text", options=("Yes", "No")),
    )),
)

# Sensitivity settings (yellow cells on the Sensitivity sheet).
SENSITIVITY_SHOCKS = (-0.20, -0.10, 0.10, 0.20)
GRID_TRAFFIC = (400, 450, 500, 550, 600)
GRID_TICKET = (13.0, 14.0, 15.0, 16.0, 17.0)

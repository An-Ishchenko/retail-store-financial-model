# Power BI dashboard spec - "Store opening: from traffic to profit"

One report page, 16:9 canvas (1280 x 720). Built on the CSV star schema in this
folder and the measures in [`measures.dax`](measures.dax).

## 1. Data model

| Table | Grain | Key columns | Notes |
|---|---|---|---|
| `fact_monthly` | month x scenario x metric | `date`, `scenario_key`, `metric_key`, `value` | Long format, 2,886 rows. Costs and outflows are negative. |
| `fact_sensitivity` | scenario x driver x shock | `scenario_key`, `driver`, `shock`, `npv`, `npv_change` | NPV with one driver changed by -20/-10/+10/+20%. |
| `dim_date` | model month | `date` | `model_month` 0-36 (0 = investment month), `model_year` 0-3, `days_in_month`. |
| `dim_scenario` | scenario | `scenario_key` | Drivers and headline KPIs from the Excel model; `sort_order` for slicers. |
| `dim_metric` | metric | `metric_key` | `metric_name`, `category`, `unit`, `sort_order`. |

Relationships (many-to-one, single direction):

```
fact_monthly[date]             -> dim_date[date]
fact_monthly[scenario_key]     -> dim_scenario[scenario_key]
fact_monthly[metric_key]       -> dim_metric[metric_key]
fact_sensitivity[scenario_key] -> dim_scenario[scenario_key]
```

Power Query: set `date` columns to Date, `value`/`npv`/`npv_change` to Decimal,
`shock` to Percentage. Sort `dim_scenario[scenario_name]` by `sort_order`,
`dim_metric[metric_name]` by `sort_order`, `dim_date[month_name]` by `month_num`.
Mark `dim_date` as a date table.

## 2. Layout

```
+--------------------------------------------------------------------------------+
| A  Title + subtitle                                   B  Scenario slicer       |
+------------+------------+------------+------------+------------+---------------+
| C1 Invest. | C2 Payback | C3 NPV     | C4 EBITDA% | C5 Break-even traffic      |
+------------+------------+------------+------------+----------------------------+
| D  Cumulative cash flow by scenario (line)        | E  Revenue & EBITDA by     |
|    all three scenarios, zero line                 |    model year (columns)    |
|                                                   +----------------------------+
|                                                   | F  NPV sensitivity         |
|                                                   |    (tornado)               |
+---------------------------------------------------+----------------------------+
| G  Footer: assumptions note + data source                                      |
+--------------------------------------------------------------------------------+
```

| # | Visual | Fields / measures | Settings |
|---|---|---|---|
| A | Text box | "Store opening: from traffic to profit" / "250 m2 home-goods store - 36-month model - fictional data, USD" | Left-aligned, 20 pt / 11 pt grey |
| B | Slicer (tiles) | `dim_scenario[scenario_name]` | Single select, default **Base** |
| C1 | Card | `SELECTEDVALUE(dim_scenario[total_investment])` | Label "Total investment", format $#,0 |
| C2 | Card | `[Payback Months]` | Label "Payback, months"; show "Not reached" when blank |
| C3 | Card | `[NPV]` | Label "NPV @ 20%", format $#,0 |
| C4 | Card | `[EBITDA %]` | Visual-level filter `dim_date[model_year] = 3`; label "EBITDA margin, Year 3" |
| C5 | Card (new card visual) | `[Break-even Traffic]`, reference label `dim_scenario[traffic_plan]` | Visual filter `model_year = 2`; label "Break-even, visitors/day (plan: N)" |
| D | Line chart | X: `dim_date[date]`; Y: `[Cumulative Cash Flow]`; Legend: `dim_scenario[scenario_name]` | Edit interactions: slicer B does **not** filter D (always show all three). Constant Y line at 0. Colours: Base #2A78D6, Pessimistic #EB6834, Optimistic #1BAF7A |
| E | Clustered column | X: `dim_date[model_year]` (filter 1-3); Y: `[Revenue]`, `[EBITDA]` | Data labels on, $#,0,K |
| F | Clustered bar | Y: `fact_sensitivity[driver]` (sort by swing); X: `SUM(fact_sensitivity[npv_change])`; Legend: `fact_sensitivity[shock]` | Visual filter `shock IN {-0.2, 0.2}`. Overlap 100%. Low shock #86B6EF, high shock #1C5CAB |
| G | Text box | "Break-even = EBITDA of zero at mature traffic. Payback excludes inventory and deposit recovered at month 36." | 9 pt grey |

Optional second page: matrix with `dim_metric[metric_name]` on rows (filter
`category IN {P&L, OPEX}`), `dim_date[model_year]` on columns, `[Metric Value]`
as values.

## 3. Expected values (to test the measures)

Base scenario, values must match the Excel model (`model/store_model.xlsx`):

| Measure | Filter | Expected |
|---|---|---|
| Revenue | Year 1 | 619,946 |
| EBITDA | Year 1 | 69,293 |
| EBITDA % | Year 3 | 14.4% |
| Net Profit | all months | 200,957 |
| Payback Months | - | 19.96 |
| NPV | - | 93,179 |
| Break-even Revenue | Year 2 | 35,327 |
| Break-even Traffic | Year 2 | 297.8 |

Other scenarios: NPV Pessimistic -51,590 (payback blank), Optimistic 225,376
(payback 12.11).

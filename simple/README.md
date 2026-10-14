# Warehouse payroll forecast model: simple version

A simpler version of the [full model](../README.md), built so the two can be compared. It uses the same synthetic data, the same 8 regions and 2 IDCs, and the same plan inputs. It has fewer moving parts and one rule for every forecast rate: **use the same month last year**.

**All data is synthetic.** The site codes (NSW-1, VIC-1, IDC-NSW and so on) are made up.

## What's different from the full model

| | Full model | Simple model |
|---|---|---|
| Forecast rates | Seasonal index x level of the last 6 months x yearly trend | Same month last year |
| Leadership | % of sales, or $ + wage increases (switch) | Share of total payroll, same month last year |
| Scenarios (upside / downside) | Yes | No |
| Yearly trend inputs | Yes | No |
| Backtest and Dashboard tabs | Yes | No |
| Public holidays and wage index | Yes | Yes |
| Cases_Forecast tab (adjustments, overrides) | Yes | Yes |
| Site codes editable, IDC served regions | Yes | Yes |
| Helper columns per site tab | ~45 | 16 |

## The core idea

Sales and payroll % come from the plan. They're given and drive everything. The model works out what that payroll buys:

```
Sales x plan payroll %          = Total payroll
Total payroll x leadership share = Leadership payroll
Total payroll - leadership      = Staff payroll
Staff payroll / cost per hour   = Hours

Sales x cases per $             = Cases     (Cases_Forecast tab, plus your adjustments)

Cases / hours                   = CPLH: the productivity the plan implies
```

The leadership share, cost per hour and cases per $ are each taken from **the same month last year**.

## How it works, step by step

An example for one region in October:

| Step | Calculation | Result |
|---|---|---|
| 1. Sales | From the plan | $150.0m |
| 2. Total payroll | $150m x 2.0% plan payroll % | $3.00m |
| 3. Leadership | $3.00m x 10.7% (leadership share of total payroll last October) | $0.32m |
| 4. Staff payroll | $3.00m - $0.32m | $2.68m |
| 5. Cost per hour | Last October's $63.50, divided by last year's wage index (1.040), times this year's (1.087). No public holidays in either October. | $66.37 |
| 6. Hours | $2.68m / $66.37 | 40,400 hours |
| 7. Cases | $150m x 0.021 cases per $ (last October) | 3.15m cases |
| 8. CPLH | 3.15m cases / 40,400 hours | 78.0 cases an hour |
| 9. Cost per case | $3.00m / 3.15m cases (total payroll) | $0.95 |

How to read it: the plan gives this site $2.68m for staff. At $66.37 an hour that buys 40,400 hours. If the site picks the same cases per $ of sales as last October, that's 3.15m cases, so it needs to run at 78.0 cases an hour. Compare that with last October's CPLH to see whether the plan is realistic. The Checks tab warns when the change is more than 10%.

## Same month last year

Each forecast month looks 12 rows up on its site tab:

- If that month is an **actual**, the rate comes from the actuals.
- If that month is a **forecast**, it uses that forecast's rate.

So Mar 2028 uses Mar 2027, which uses Mar 2026 actuals. When you add a year of actuals and move **Last actual month** forward, Mar 2027 becomes an actual, and Mar 2028 automatically picks it up. The model "actualises" itself one year at a time.

**Why:** it's the easiest rule to explain and check. It keeps each month's own pattern (December is busier, February is quieter) without a separate seasonal calculation. The catch is that one unusual month last year carries straight into the forecast. If that happens, use the Cases_Forecast overrides.

## Leadership

Leadership gets different pay increases from staff, so it isn't put through the staff wage index. Instead the model takes leadership's **share of total payroll** in the same month last year:

```
Leadership = total payroll x (last year's leadership payroll / last year's total payroll)
```

Last year's share already reflects how leadership pay moved relative to staff pay.

## Public holidays and wage increases (cost per hour)

Cost per hour moves for two reasons that have nothing to do with how a site runs:

- **Public holidays:** penalty rates push up the average cost per hour.
- **Wage increases:** every hour costs more from the date they start.

So last year's cost per hour isn't copied as it is. The model:

1. **Removes last year's effects:** it divides by last year's public-holiday factor and wage index.
2. **Adds this year's back:** it multiplies by this month's public-holiday factor and wage index.

So if Easter falls in April last year and March this year, or a 4.5% increase starts in January, the cost lands in the right month.

**Public-holiday factor**, for example a 30-day month with 2 public holidays, where half the normal hours are worked on a holiday at 250% pay:

```
(28 normal days + 2 holidays x 0.5 of the hours x 2.5 pay)  /  (28 + 2 x 0.5)  =  30.5 / 29  =  1.052
```

The share of hours worked and the pay loading are on the Inputs tab. Holidays are listed by state on the Holidays tab, and the Calendar tab counts them per month.

**Wage index:** each increase multiplies the one before. With 4.0% in Jan 2026 and 4.5% in Jan 2027, the index is 1.040 for 2026 and 1.087 from Jan 2027. Keep past increases in the wage table too, so last year's cost per hour can be adjusted correctly.

## Cases forecast

The Cases_Forecast tab has months down the side and one column per site, in four blocks:

1. **Run rate**: volume driver sales x cases per $ in the same month last year. Actual months show actual cases.
2. **Adjustments**: type extra cases (or a negative number) for a site and month, e.g. a new store.
3. **Override**: type the cases you want. It replaces the run rate and the adjustment.
4. **Cases used**: the override if there is one, otherwise run rate + adjustment. This feeds each site tab.

Adjustments and overrides are separate yellow input cells, so typing in them never overwrites a formula. Clear a cell to go back to the model. Anything typed in an actual month is ignored, and Checks warns about it.

## IDCs

The two IDCs have no sales of their own:

- **Payroll** uses the host region's sales and the IDC's plan payroll %: IDC-NSW uses NSW-1, IDC-VIC uses VIC-1.
- **Cases** follow the combined sales of the regions they ship to: IDC-NSW serves NSW-1 + NSW-2, IDC-VIC serves VIC-1 + VIC-2.

The National tab takes sales from the 8 regions only, so the host's sales aren't counted twice.

## Using your own site codes

Type your codes in the **Code** column of the Sites table on the Inputs tab. The sample data, Checks, National and each site tab's title all follow. Renaming the tabs is optional, because Excel updates formulas when a tab is renamed. Paste real actuals and plans with the same codes.

## Updating each month or year

1. Paste the new actuals into Data_Actuals and the plan into Plan_Inputs (same columns, one row per site per month).
2. Move **Last actual month** on the Inputs tab forward.
3. Review Cases_Forecast, and clear adjustments or overrides for months that are now actuals.
4. Check that the Checks tab says **ALL CHECKS OK**.

The model needs at least 12 months of actuals, so every forecast month has a "last year" to look at.

## Workbook tabs

- **README**: the logic in short
- **Inputs**: last actual month, CPLH warning tolerance, wage increases, public-holiday assumptions, site codes
- **Data_Actuals / Plan_Inputs**: actuals and plan, one row per site per month
- **Cases_Forecast**: run rate, adjustments, overrides and cases used
- **CLR**: placeholder for cost-saving initiatives (not linked)
- **Holidays / Calendar**: public holidays by state, PH cost factor, wage index
- **Checks**: data coverage, reconciliation and warnings
- **National**: all sites added up, plus site-by-year tables (payroll, cases, payroll %, CPLH, cost per case)
- **One tab per site**: drivers and metrics on the left, 16 helper columns in grey on the right

Colours: yellow cells with blue text are inputs. Blue-shaded rows are forecast months.

## Build

From the repo root:

```bash
python3 -m venv .venv
.venv/bin/pip install openpyxl
.venv/bin/python simple/build_model.py
```

This outputs `simple/warehouse_payroll_forecast_simple.xlsx`. The workbook recalculates when it opens.

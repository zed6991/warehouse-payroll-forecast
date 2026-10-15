# Warehouse payroll forecast model: CPLH-driven version

A version of the [simple model](../simple/README.md) with the chain turned around. In the simple model, the plan's payroll % is given and the model works out the CPLH that payroll implies. Here **CPLH and cases are the inputs**, and the model works out the payroll, and the payroll % of sales, that they need.

**All data is synthetic.** The site codes (NSW-1, VIC-1, IDC-NSW and so on) are made up.

## What's different from the simple model

| | Simple model | CPLH-driven model |
|---|---|---|
| Inputs that drive payroll | Plan sales x plan payroll % | Cases / CPLH x cost per hour |
| CPLH | Output (cases / hours) | Input: same month last year, plus adjustments or overrides on the CPLH_Forecast tab |
| Payroll % of sales | Input from the plan | Output (total payroll / sales) |
| Leadership share of total payroll | Same month last year | Last actual month, for every forecast month |
| Plan_Inputs | Sales and payroll % | Sales only |
| Cases_Forecast, cost per hour, public holidays, wage index | Same | Same |

## The core idea

```
Cases / CPLH                                 = Hours
Hours x cost per hour                        = Staff payroll
Staff payroll / (1 - leadership share)       = Total payroll
Total payroll - staff payroll                = Leadership payroll

Total payroll / sales                        = Payroll % of sales: what the productivity costs
```

## How it works, step by step

An example for one region in October:

| Step | Calculation | Result |
|---|---|---|
| 1. CPLH | Last October's CPLH, from the CPLH_Forecast tab | 78.0 cases an hour |
| 2. Cases | $150m x 0.021 cases per $ (last October), from the Cases_Forecast tab | 3.15m cases |
| 3. Hours | 3.15m cases / 78.0 | 40,385 hours |
| 4. Cost per hour | Last October's $63.50, divided by last year's wage index (1.040), times this year's (1.087) | $66.37 |
| 5. Staff payroll | 40,385 hours x $66.37 | $2.68m |
| 6. Sales | From the plan | $150.0m |
| 7. Staff payroll % | $2.68m / $150m | 1.79% |
| 8. Total payroll | $2.68m / (1 - 10.7% leadership share in the last actual month) | $3.00m |
| 9. Leadership payroll | $3.00m - $2.68m | $0.32m |
| 10. Payroll % of sales | $3.00m / $150m | 2.00% |

How to read it: if the site picks 3.15m cases at last October's 78 cases an hour, it needs 40,385 hours. At $66.37 an hour that's $2.68m of staff payroll, or $3.00m including leadership, which is 2.0% of sales. Type a higher CPLH (say a productivity target) and payroll % comes down.

## CPLH forecast

The CPLH_Forecast tab works like Cases_Forecast: months down the side, one column per site, in four blocks:

1. **Run rate**: CPLH in the same month last year. Actual months show actual CPLH.
2. **Adjustments**: type a change in cases per hour, e.g. +2 for a new pick system.
3. **Override**: type the CPLH you want, e.g. a target. It replaces the run rate and the adjustment.
4. **CPLH used**: the override if there is one, otherwise run rate + adjustment. This feeds each site tab.

The run rate chains back to actuals only: Oct-27 uses Oct-26's run rate, which uses Oct-25 actuals. An adjustment or override in Oct-26 is **not** carried into Oct-27. Type it again if it should continue. The same is true for cases.

The Checks tab warns when a forecast month's CPLH (after adjustments and overrides) moves more than the tolerance on the Inputs tab vs the same month last year.

## Leadership

Leadership is worked out as a share of total payroll, using **the last actual month's** share for every forecast month (shown at the top of each site tab, cell K2). So if staff payroll goes up, leadership goes up in proportion. When you move **Last actual month** forward, the share updates.

This is a flat share, so it doesn't pick up the November and December leadership peaks that the same-month-last-year rule in the simple model keeps.

## IDCs

As in the simple model: IDC cases follow the combined sales of the regions they serve, and IDC payroll % is shown against the host region's sales (IDC-NSW against NSW-1, IDC-VIC against VIC-1).

## Updating each month or year

1. Paste the new actuals into Data_Actuals and the plan sales into Plan_Inputs.
2. Move **Last actual month** on the Inputs tab forward.
3. Review Cases_Forecast and CPLH_Forecast, and clear adjustments or overrides for months that are now actuals.
4. Check that the Checks tab says **ALL CHECKS OK**.

## Build

From the repo root:

```bash
.venv/bin/python cplh-driven/build_model.py
```

This outputs `cplh-driven/warehouse_payroll_forecast_cplh.xlsx`. The workbook recalculates when it opens.

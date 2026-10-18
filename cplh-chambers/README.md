# Warehouse payroll forecast model: CPLH-driven, with chambers

A copy of the [CPLH-driven model](../cplh-driven/README.md) with three additions:

1. **Plan CPLH by site and month** is an input on Plan_Inputs.
2. **CPLH_Breakdown** splits one national CPLH target per year into site CPLH by month, based on each site's actual CPLH.
3. **A chamber split**: each site's cases, hours and payroll are split into ambient, chiller and freezer.

**All data is synthetic.** The site codes and the chamber mix are made up.

## The core idea (unchanged)

```
Cases / CPLH                                 = Hours
Hours x cost per hour                        = Staff payroll
Staff payroll / (1 - leadership share)       = Total payroll
Total payroll - staff payroll                = Leadership payroll

Total payroll / sales                        = Payroll % of sales
```

## Plan CPLH

Plan_Inputs now has a **Plan CPLH** column next to sales, with one row per site per month. The CPLH_Forecast tab uses it like this:

1. **Plan CPLH**: the plan for that site and month. If it's blank, it falls back to CPLH in the same month last year.
2. **Adjustments** and 3. **Override**: as before.
4. **CPLH used**: feeds hours, and from there payroll, on each site tab.

The Checks tab warns when a site is missing a plan CPLH for a forecast month (it still runs on last year's CPLH).

**Plan CPLH source** on the Inputs tab switches where the plan comes from:
- **Plan_Inputs** (the default): what you typed.
- **National breakdown**: the suggested CPLH from the CPLH_Breakdown tab.

## CPLH_Breakdown: national target to sites

Type a national CPLH target for a year (cases / hours over the included sites). The tab scales every included site's **base CPLH** by one factor for that year:

- **Base CPLH** is the site's actual CPLH, or for forecast months, the actual CPLH in the same month last year.
- Each site keeps its own level and seasonal shape. Only the factor changes.

The factor is worked out so the full year, actual months included, lands exactly on the target:

```
target = FY cases / (actual hours + forecast cases / (base CPLH x factor))

factor = forecast hours at base CPLH / (FY cases / target - actual hours)
```

Example: FY2027 base national CPLH is 77.7 and the target is 80.9, so the factor is 1.042. Every included site's CPLH goes up 4.2% in every FY2027 month. A site at 79.5 in Oct-26 gets 82.8 in Oct-27.

- **Include** row: pick which sites count towards the national number. The default is the 8 regions. IDCs run at much higher CPLH and would skew the mix. Excluded sites keep their base CPLH.
- **Blank target**: no change (factor 1).
- **Part-actual year**: the target applies to the whole year, so the forecast months make up the difference. If the actual months already use more hours than the target allows, the factor shows #N/A and the Checks tab flags it.
- **Using it**: either copy the paste column on the far right into Plan_Inputs column D (it's in the same row order) and paste as values, or set Plan CPLH source to National breakdown.

## Chambers

Each site's split uses one case mix: its share of cases in ambient, chiller and freezer, from the **Chamber Weights** tab.

- Paste cases by chamber for each site into the yellow cells (e.g. the last 3 months, if the mix has changed). The % table below them works out the shares.
- The % table sits at `B24:K26`, with site codes in row 23 and chamber names in `A24:A26`, the same addresses as the live workbook. The lookup on each site tab is the same `INDEX/MATCH` formula.
- Every chamber uses the site's CPLH and cost per hour, so cases, hours, staff, leadership and total payroll are all split by the case mix.
- The split covers actual months too, using the same mix, because the actuals have no chamber detail.

Where to see it:
- **Each site tab**, to the right of the helper columns (from column AJ): one block per chamber with cases, hours, staff, leadership and total payroll for every month, FY totals in rows 6-9, and the mix in row 11.
- **National**: the same blocks, added up from the sites. Row 11 shows each chamber's share of national cases.
- **Chamber_Summary**: pick a FY to see cases, hours, staff and total payroll by site and chamber.

The Checks tab confirms each site's weights add to 100%, and that national chamber totals add back to national cases, hours and total payroll.

## Updating each month or year

1. Paste the new actuals into Data_Actuals, and the plan sales and plan CPLH into Plan_Inputs.
2. Paste the latest case mix into Chamber Weights.
3. Move **Last actual month** on the Inputs tab forward.
4. Review Cases_Forecast and CPLH_Forecast. Clear adjustments or overrides for months that are now actuals.
5. Check that the Checks tab says **ALL CHECKS OK**.

## Build

From the repo root:

```bash
.venv/bin/python cplh-chambers/build_model.py
```

This outputs `cplh-chambers/warehouse_payroll_forecast_cplh_chambers.xlsx`. The workbook recalculates when it opens.

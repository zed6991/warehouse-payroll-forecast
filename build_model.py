"""Build the warehouse payroll forecast model (synthetic data) as an Excel workbook.

Run:  .venv/bin/python build_model.py
Output: warehouse_payroll_forecast.xlsx
"""
import calendar
import math
import random
from datetime import date, timedelta

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

OUT = "warehouse_payroll_forecast.xlsx"
# bounded ranges (whole-column refs are slow); room for pasted data up to these rows
DATA_MAX, HOL_MAX = 5000, 1000
DR = lambda col: f"Data_Actuals!${col}$2:${col}${DATA_MAX}"
PR = lambda col: f"Plan_Inputs!${col}$2:${col}${DATA_MAX}"
HR = lambda col: f"Holidays!${col}$2:${col}${HOL_MAX}"
START = date(2025, 1, 1)
END = date(2028, 12, 1)
LAST_ACTUAL = date(2026, 9, 1)
MONTHS = []
d = START
while d <= END:
    MONTHS.append(d)
    d = date(d.year + (d.month == 12), d.month % 12 + 1, 1)
N = len(MONTHS)  # 48
FIRST_ROW = 15
LAST_ROW = FIRST_ROW + N - 1  # 62
YEARS = [2025, 2026, 2027, 2028]

# code, type, state, host region (IDCs borrow host sales)
SITES = [
    ("NSW-1", "Region", "NSW", None),
    ("VIC-1", "Region", "VIC", None),
    ("QLD-1", "Region", "QLD", None),
    ("NSW-2", "Region", "NSW", None),
    ("VIC-2", "Region", "VIC", None),
    ("QLD-2", "Region", "QLD", None),
    ("SA-1", "Region", "SA", None),
    ("WA-1", "Region", "WA", None),
    ("IDC-NSW", "IDC", "NSW", "NSW-1"),
    ("IDC-VIC", "IDC", "VIC", "VIC-1"),
]
REGIONS = [s[0] for s in SITES if s[1] == "Region"]
STATES = ["NSW", "VIC", "QLD", "SA", "WA"]

WAGE_TABLE = [(date(2026, 1, 1), 0.040), (date(2027, 1, 1), 0.045), (date(2028, 1, 1), 0.045)]
PH_ROSTER = 0.5   # share of a normal day's hours worked on a public holiday
PH_LOADING = 1.5  # extra pay on a public holiday (1.5 = paid 250%)


# ---------------------------------------------------------------- holidays
def easter(y):
    a, b, c = y % 19, y // 100, y % 100
    dd, e = b // 4, b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - dd - g + 15) % 30
    i, k = c // 4, c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = (h + l - 7 * m + 114) % 31 + 1
    return date(y, month, day)


def nth_weekday(y, m, wd, n):
    """n-th weekday (0=Mon) of a month; n=-1 for the last."""
    if n > 0:
        d0 = date(y, m, 1)
        return d0 + timedelta(days=(wd - d0.weekday()) % 7 + 7 * (n - 1))
    d0 = date(y, m, calendar.monthrange(y, m)[1])
    return d0 - timedelta(days=(d0.weekday() - wd) % 7)


def next_monday(d0):
    return d0 + timedelta(days=(7 - d0.weekday()) % 7)


def holidays():
    rows = []
    for y in YEARS:
        ny = date(y, 1, 1)
        aus = date(y, 1, 26)
        gf = easter(y) - timedelta(days=2)
        xmas = date(y, 12, 25)
        anzac = date(y, 4, 25)
        for st in STATES:
            h = [(ny, "New Year's Day")]
            if ny.weekday() >= 5:
                h.append((next_monday(ny), "New Year's Day (additional)"))
            h.append((next_monday(aus) if aus.weekday() >= 5 else aus, "Australia Day"))
            h.append((gf, "Good Friday"))
            if st != "WA":
                h.append((gf + timedelta(days=1), "Easter Saturday"))
            h.append((gf + timedelta(days=2), "Easter Sunday"))
            h.append((gf + timedelta(days=3), "Easter Monday"))
            h.append((anzac, "Anzac Day"))
            if (st == "WA" and anzac.weekday() >= 5) or (st == "SA" and anzac.weekday() == 6):
                h.append((next_monday(anzac), "Anzac Day (additional)"))
            if st in ("NSW", "VIC", "SA"):
                h.append((nth_weekday(y, 6, 0, 2), "King's Birthday"))
            elif st == "QLD":
                h.append((nth_weekday(y, 10, 0, 1), "King's Birthday"))
            else:
                h.append((nth_weekday(y, 9, 0, -1), "King's Birthday"))
            labour = {"NSW": (10, 1), "SA": (10, 1), "VIC": (3, 2), "QLD": (5, 1), "WA": (3, 1)}[st]
            h.append((nth_weekday(y, labour[0], 0, labour[1]), "Labour Day"))
            if st == "VIC":
                h.append((nth_weekday(y, 11, 1, 1), "Melbourne Cup"))
                h.append((nth_weekday(y, 9, 5, -1) - timedelta(days=1), "AFL Grand Final Friday (check date)"))
            if st == "SA":
                h.append((nth_weekday(y, 3, 0, 2), "Adelaide Cup"))
            if st == "WA":
                h.append((nth_weekday(y, 6, 0, 1), "WA Day"))
            h.append((xmas, "Christmas Day"))
            h.append((date(y, 12, 26), "Proclamation Day" if st == "SA" else "Boxing Day"))
            if xmas.weekday() == 5:
                h += [(date(y, 12, 27), "Christmas (additional)"), (date(y, 12, 28), "Boxing Day (additional)")]
            elif xmas.weekday() == 6:
                h.append((date(y, 12, 27), "Christmas (additional)"))
            elif xmas.weekday() == 4:
                h.append((date(y, 12, 28), "Boxing Day (additional)"))
            rows += [(dt, name, st) for dt, name in h]
    rows.sort(key=lambda r: (r[2], r[0]))
    return rows


HOLIDAYS = holidays()


def ph_factor(month, state):
    days = calendar.monthrange(month.year, month.month)[1]
    h = sum(1 for dt, _, st in HOLIDAYS if st == state and dt.year == month.year and dt.month == month.month)
    return (days - h + h * PH_ROSTER * (1 + PH_LOADING)) / (days - h + h * PH_ROSTER)


def wage_index(month):
    idx = 1.0
    for eff, pct in WAGE_TABLE:
        if eff <= month:
            idx *= 1 + pct
    return idx


# ---------------------------------------------------------- synthetic data
SEASON_SALES = {1: .97, 2: .90, 3: 1.0, 4: .99, 5: .98, 6: .96, 7: .98, 8: .97, 9: .96, 10: 1.0, 11: 1.03, 12: 1.25}
SEASON_CPLH = {1: .97, 2: .99, 3: 1.0, 4: .99, 5: 1.0, 6: 1.0, 7: 1.0, 8: 1.0, 9: 1.0, 10: 1.01, 11: 1.03, 12: 1.06}
SEASON_LEAD = {11: 1.05, 12: 1.15}
STATE_CPH = {"NSW": 66.0, "VIC": 65.0, "QLD": 63.0, "SA": 61.0, "WA": 69.0}
# monthly sales $m (2025), cases per $ sales, CPLH, leadership $m per month
PARAMS = {
    "NSW-1": (150, .0210, 78, .32), "VIC-1": (140, .0215, 80, .30), "QLD-1": (115, .0220, 74, .26),
    "NSW-2": (125, .0212, 76, .28), "VIC-2": (135, .0208, 79, .30), "QLD-2": (120, .0218, 75, .27),
    "SA-1": (80, .0225, 72, .21), "WA-1": (105, .0223, 73, .25),
    "IDC-NSW": (None, .0060, 110, .15), "IDC-VIC": (None, .0055, 105, .14),
}


def synth():
    rng = random.Random(42)
    data = {}
    for code, typ, state, host in SITES:
        base_sales, cps0, cplh0, lead0 = PARAMS[code]
        for t, m in enumerate(MONTHS):
            if typ == "Region":
                sales = base_sales * 1e6 * SEASON_SALES[m.month] * 1.035 ** (t / 12) * (1 + rng.gauss(0, .015))
            else:
                sales = data[(host, m)]["sales"]
            cps = cps0 * 0.985 ** (t / 12) * (0.98 if m.month == 12 else 1) * (1 + rng.gauss(0, .01))
            cases = sales * cps
            cplh = cplh0 * SEASON_CPLH[m.month] * 1.01 ** (t / 12) * (1 + rng.gauss(0, .015))
            hours = cases / cplh
            cph = STATE_CPH[state] * (1.06 if typ == "IDC" else 1) * wage_index(m) * ph_factor(m, state) * (1 + rng.gauss(0, .008))
            staff = hours * cph
            lead = lead0 * 1e6 * wage_index(m) * SEASON_LEAD.get(m.month, 1) * (1 + rng.gauss(0, .02))
            data[(code, m)] = dict(sales=sales, cases=cases, hours=hours, staff=staff, lead=lead)
    return data


DATA = synth()

# ------------------------------------------------------------------ styles
NAVY = PatternFill("solid", fgColor="1F3864")
LIGHT = PatternFill("solid", fgColor="DDEBF7")
INPUT = PatternFill("solid", fgColor="FFF2CC")
GREY = PatternFill("solid", fgColor="F2F2F2")
WHITE_B = Font(bold=True, color="FFFFFF")
BOLD = Font(bold=True)
BLUE = Font(color="0000FF")
HELP = Font(color="808080", italic=True)
THIN = Side(style="thin", color="1F3864")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

F_M = '#,##0.0,,'        # $ or units shown in millions
F_CASES = '#,##0.00,,'
F_INT = '#,##0'
F_1 = '0.0'
F_CPH = '$#,##0.00'
F_CPC = '$#,##0.000'
F_PCT = '0.00%'
F_DATE = 'mmm-yy'
F_RATIO = '0.00000'

wb = Workbook()


def header(ws, row, col, text, fill=NAVY, font=WHITE_B, width=None):
    c = ws.cell(row=row, column=col, value=text)
    c.fill, c.font, c.alignment = fill, font, CENTER
    if width:
        ws.column_dimensions[get_column_letter(col)].width = width
    return c


def name(nm, ref):
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)


# ------------------------------------------------------------------ README
ws = wb.active
ws.title = "README"
readme = [
    ("Warehouse payroll forecast model", BOLD),
    ("Synthetic data only. Monthly, Jan 2025 to Dec 2028.", None),
    ("", None),
    ("HOW TO UPDATE ACTUALS", BOLD),
    ("1. Paste the new actuals over the rows in Data_Actuals (keep the header and column order). One row per site per month.", None),
    ("   Months must be the 1st of the month. IDC rows leave Sales blank.", None),
    ("2. Set 'Last actual month' on the Inputs tab. Every tab switches actual/forecast from that date.", None),
    ("3. Paste the plan (sales forecast + payroll % of sales) into Plan_Inputs, same layout.", None),
    ("4. Check the Checks tab says ALL CHECKS OK.", None),
    ("", None),
    ("FORECAST LOGIC (per site, per month)", BOLD),
    ("Sales = plan sales x (1 + scenario sales lever). IDCs use their host region's sales (IDC-NSW = NSW-1, IDC-VIC = VIC-1).", None),
    ("Total payroll = sales x plan payroll % (the driver).", None),
    ("Leadership = sales x leadership % run rate.  Staff payroll = total payroll - leadership.", None),
    ("Cost per hour = run rate x public-holiday factor x wage index.  Hours = staff payroll / cost per hour.", None),
    ("Cases = sales x cases-per-$ run rate x (1 + scenario cases lever).", None),
    ("CPLH = cases / hours.  Cost per case = staff payroll / cases (excludes leadership).", None),
    ("", None),
    ("RUN RATES (cases per $ sales, cost per hour, leadership %)", BOLD),
    ("Seasonal index = each calendar month vs the average over the last 12 actual months.", None),
    ("Level = average of the last 6 actual months / average seasonal index of those months.", None),
    ("Forecast = level x seasonal index, so the last 6 months set the level and last year sets the shape.", None),
    ("Cost per hour is first stripped of public-holiday and wage effects, then they are added back for each forecast month.", None),
    ("So Easter moving from April to March, or an extra holiday, changes cost per hour in that month.", None),
    ("", None),
    ("NATIONAL", BOLD),
    ("Sales = 8 regions only (IDC sales are a copy of NSW-1/VIC-1 and are excluded). Cases, hours and payroll = all 10 sites.", None),
    ("", None),
    ("COLOURS", BOLD),
    ("Yellow cells with blue text are inputs. Blue-shaded rows are forecast months. Grey columns are helper calculations.", None),
    ("Public holidays are generated by rule; check against the official state lists (AFL Grand Final Friday in particular).", None),
]
for i, (txt, font) in enumerate(readme, 1):
    c = ws.cell(row=i, column=1, value=txt)
    if font:
        c.font = font
ws.column_dimensions["A"].width = 120

# ------------------------------------------------------------------ Inputs
ws = wb.create_sheet("Inputs")
ws.column_dimensions["A"].width = 34
for col in "BCDEF":
    ws.column_dimensions[col].width = 14
ws["A1"] = "Inputs"
ws["A1"].font = Font(bold=True, size=14)

ws["A3"] = "Last actual month"
ws["B3"] = LAST_ACTUAL
ws["B3"].number_format = F_DATE
ws["A4"] = "Run-rate months (level)"
ws["B4"] = 6
ws["A5"] = "Seasonality window (months)"
ws["B5"] = 12
ws["A6"] = "Model status"
ws["B6"] = "=Checks!B3"
ws["B6"].font = BOLD
for r in (3, 4, 5):
    ws.cell(row=r, column=2).fill, ws.cell(row=r, column=2).font = INPUT, BLUE
name("LastActual", "Inputs!$B$3")
name("RunRateMonths", "Inputs!$B$4")
name("SeasonMonths", "Inputs!$B$5")

ws["A8"] = "Scenario"
ws["A8"].font = BOLD
ws["A9"] = "Selected scenario"
ws["B9"] = "Base"
ws["B9"].fill, ws["B9"].font = INPUT, BLUE
dv = DataValidation(type="list", formula1='"Base,Upside,Downside"', allow_blank=False)
ws.add_data_validation(dv)
dv.add("B9")
for c, t in zip("ABC", ["Scenario", "Sales lever", "Cases lever"]):
    header(ws, 11, " ABC".index(c), t)
for i, (sc, s, cs) in enumerate([("Base", 0, 0), ("Upside", .03, .02), ("Downside", -.03, -.02)]):
    r = 12 + i
    ws.cell(row=r, column=1, value=sc)
    for col, v in ((2, s), (3, cs)):
        c = ws.cell(row=r, column=col, value=v)
        c.number_format, c.fill, c.font = '0.0%', INPUT, BLUE
ws["A16"] = "Sales lever (active)"
ws["B16"] = "=INDEX($B$12:$B$14,MATCH($B$9,$A$12:$A$14,0))"
ws["A17"] = "Cases lever (active)"
ws["B17"] = "=INDEX($C$12:$C$14,MATCH($B$9,$A$12:$A$14,0))"
ws["B16"].number_format = ws["B17"].number_format = '0.0%'
name("SalesAdj", "Inputs!$B$16")
name("CasesAdj", "Inputs!$B$17")

ws["A19"] = "Wage increases (applied to cost per hour)"
ws["A19"].font = BOLD
header(ws, 20, 1, "Effective month")
header(ws, 20, 2, "Increase")
for i in range(8):
    r = 21 + i
    a, b = ws.cell(row=r, column=1), ws.cell(row=r, column=2)
    if i < len(WAGE_TABLE):
        a.value, b.value = WAGE_TABLE[i]
    a.number_format, b.number_format = F_DATE, '0.0%'
    a.fill = b.fill = INPUT
    a.font = b.font = BLUE
ws["C21"] = "History row: lets the model strip past increases out of actual cost per hour. Replace with your real ones."
ws["C21"].font = HELP
name("WageDates", "Inputs!$A$21:$A$28")
name("WageRates", "Inputs!$B$21:$B$28")

ws["A30"] = "Public holidays"
ws["A30"].font = BOLD
ws["A31"] = "Share of normal hours worked on a PH"
ws["B31"] = PH_ROSTER
ws["A32"] = "PH pay loading (1.5 = paid 250%)"
ws["B32"] = PH_LOADING
for r in (31, 32):
    ws.cell(row=r, column=2).fill, ws.cell(row=r, column=2).font = INPUT, BLUE
name("PH_Roster", "Inputs!$B$31")
name("PH_Loading", "Inputs!$B$32")

ws["A34"] = "Sites"
ws["A34"].font = BOLD
for j, t in enumerate(["Code", "Type", "State", "Sales source", "In national sales?"], 1):
    header(ws, 35, j, t)
for i, (code, typ, st, host) in enumerate(SITES):
    r = 36 + i
    vals = [code, typ, st, host or code, "No" if typ == "IDC" else "Yes"]
    for j, v in enumerate(vals, 1):
        ws.cell(row=r, column=j, value=v)
name("SiteCodes", f"Inputs!$A$36:$A${35 + len(SITES)}")

# ------------------------------------------------------------ Data_Actuals
ws = wb.create_sheet("Data_Actuals")
cols = ["Month", "Site", "Sales $", "Cases picked", "WH hours worked", "Staff payroll $", "Leadership payroll $"]
for j, t in enumerate(cols, 1):
    header(ws, 1, j, t, width=16)
r = 2
for m in MONTHS:
    if m > LAST_ACTUAL:
        break
    for code, typ, _, _ in SITES:
        v = DATA[(code, m)]
        row = [m, code, None if typ == "IDC" else round(v["sales"]), round(v["cases"]), round(v["hours"]),
               round(v["staff"]), round(v["lead"])]
        for j, val in enumerate(row, 1):
            c = ws.cell(row=r, column=j, value=val)
            c.number_format = F_DATE if j == 1 else F_INT
        r += 1
ws.freeze_panes = "A2"

# ------------------------------------------------------------- Plan_Inputs
ws = wb.create_sheet("Plan_Inputs")
for j, t in enumerate(["Month", "Site", "Sales forecast $", "Payroll % of sales (total)"], 1):
    header(ws, 1, j, t, width=22)
ws["F1"] = "Stand-in for the long-range plan. IDC payroll % is a % of the host region's sales (IDC-NSW: NSW-1, IDC-VIC: VIC-1)."
ws["F1"].font = HELP
r = 2
for m in MONTHS:
    if m.year < 2026:
        continue
    for code, typ, _, host in SITES:
        v = DATA[(code, m)]
        pct = round((v["staff"] + v["lead"]) / v["sales"], 4)
        row = [m, code, None if typ == "IDC" else round(v["sales"], -3), pct]
        for j, val in enumerate(row, 1):
            c = ws.cell(row=r, column=j, value=val)
            c.number_format = [F_DATE, None, F_INT, F_PCT][j - 1] or "General"
        r += 1
ws.freeze_panes = "A2"

# ---------------------------------------------------------------- Holidays
ws = wb.create_sheet("Holidays")
for j, t in enumerate(["Date", "Holiday", "State"], 1):
    header(ws, 1, j, t, width=[12, 38, 8][j - 1])
for i, (dt, nm, st) in enumerate(HOLIDAYS, 2):
    ws.cell(row=i, column=1, value=dt).number_format = "dd-mmm-yy"
    ws.cell(row=i, column=2, value=nm)
    ws.cell(row=i, column=3, value=st)
ws["E1"] = "Generated by rule. Add, remove or move rows freely; the Calendar tab counts them by state and month."
ws["E1"].font = HELP
ws.freeze_panes = "A2"

# ---------------------------------------------------------------- Calendar
ws = wb.create_sheet("Calendar")
ws["A1"] = "Calendar: public holidays, PH cost factor, wage index"
ws["A1"].font = BOLD
ws["A2"] = "PH factor = average cost per hour in the month vs a month with no public holidays."
ws["A2"].font = HELP
heads = ["Month", "Days"] + [f"PH {s}" for s in STATES] + STATES + ["Wage index"]
for j, t in enumerate(heads, 1):
    header(ws, 4, j, t, width=11)
ws.cell(row=3, column=3, value="Public holidays (count)").font = BOLD
ws.cell(row=3, column=8, value="PH cost factor").font = BOLD
for i, m in enumerate(MONTHS):
    r = 5 + i
    ws.cell(row=r, column=1, value=m).number_format = F_DATE
    ws.cell(row=r, column=2, value=f"=DAY(EOMONTH(A{r},0))")
    for k, st in enumerate(STATES):
        cnt = 3 + k
        ws.cell(row=r, column=cnt, value=(
            f'=COUNTIFS({HR("A")},">="&$A{r},{HR("A")},"<="&EOMONTH($A{r},0),{HR("C")},"{st}")'))
        h = get_column_letter(cnt)
        c = ws.cell(row=r, column=8 + k, value=(
            f"=($B{r}-{h}{r}+{h}{r}*PH_Roster*(1+PH_Loading))/($B{r}-{h}{r}+{h}{r}*PH_Roster)"))
        c.number_format = '0.0000'
    # product of (1 + increase) for every increase effective on or before this month
    c = ws.cell(row=r, column=13, value=f"=EXP(SUMPRODUCT((WageDates<=$A{r})*(WageDates<>\"\")*LN(1+WageRates)))")
    c.number_format = '0.0000'
ws.freeze_panes = "B5"

# -------------------------------------------------------------- site tabs
SITE_COLS = {
    "A": ("Month", F_DATE, 9), "B": ("Year", "0", 6), "C": ("Mth", "0", 5), "D": ("Status", None, 9),
    "E": ("Sales", F_M, 10), "F": ("Cases picked", F_CASES, 10), "G": ("WH hours worked", F_INT, 11),
    "H": ("Staff payroll", F_M, 10), "I": ("Leadership payroll", F_M, 11), "J": ("Total payroll", F_M, 10),
    "K": ("CPLH (cases / hr)", F_1, 10), "L": ("Cost per hour", F_CPH, 10), "M": ("Cost per case", F_CPC, 10),
    "N": ("Staff payroll % sales", F_PCT, 10), "O": ("Leadership % sales", F_PCT, 10), "P": ("Total payroll % sales", F_PCT, 10),
    "R": ("Act sales", F_INT, 11), "S": ("Act cases", F_INT, 11), "T": ("Act hours", F_INT, 10),
    "U": ("Act staff $", F_INT, 11), "V": ("Act leadership $", F_INT, 11),
    "W": ("Act cases per $", F_RATIO, 10), "X": ("Act CPH (ex PH & wage)", F_CPH, 11), "Y": ("Act leadership %", F_PCT, 10),
    "Z": ("PH factor", '0.0000', 8), "AA": ("Wage index", '0.0000', 8),
    "AB": ("Plan sales", F_INT, 12), "AC": ("Plan payroll %", F_PCT, 9),
    "AD": ("Season idx cases/$", '0.000', 9), "AE": ("Season idx CPH", '0.000', 9), "AF": ("Season idx lead %", '0.000', 9),
    "AG": ("Fcst cases per $", F_RATIO, 10), "AH": ("Fcst CPH", F_CPH, 9), "AI": ("Fcst leadership %", F_PCT, 9),
}
UNITS = {"E": "$m", "F": "m", "G": "#", "H": "$m", "I": "$m", "J": "$m", "K": "#", "L": "$", "M": "$",
         "N": "%", "O": "%", "P": "%"}
RNG = lambda col: f"${col}${FIRST_ROW}:${col}${LAST_ROW}"
ACT_WIN = lambda months: f'{RNG("A")},">"&EDATE(LastActual,-{months}),{RNG("A")},"<="&LastActual'


def data_sum(col, r):
    return f"SUMIFS({DR(col)},{DR('A')},$A{r},{DR('B')},$B$2)"


def style_site_sheet(ws):
    for col, (title, fmt, width) in SITE_COLS.items():
        ws.column_dimensions[col].width = width
        c = header(ws, 13, ws[f"{col}1"].column, title)
        if col >= "R" or len(col) == 2:
            c.fill = PatternFill("solid", fgColor="7F7F7F")
        if col in UNITS:
            u = ws[f"{col}14"]
            u.value, u.font, u.alignment, u.fill = UNITS[col], Font(italic=True), CENTER, LIGHT
    ws.column_dimensions["Q"].width = 2
    ws.row_dimensions[13].height = 45
    for col, text, span in (("E", "Drivers", "E12:J12"), ("K", "Metrics", "K12:P12"),
                            ("R", "Helper calculations (actuals, run rates, plan)", "R12:AI12")):
        ws.merge_cells(span)
        header(ws, 12, ws[f"{col}1"].column, text)
    ws.freeze_panes = f"E{FIRST_ROW}"
    ws.conditional_formatting.add(
        f"A{FIRST_ROW}:P{LAST_ROW}", FormulaRule(formula=[f'$D{FIRST_ROW}="Forecast"'], fill=LIGHT))
    for r in range(FIRST_ROW, LAST_ROW + 1):
        for col, (_, fmt, _) in SITE_COLS.items():
            c = ws[f"{col}{r}"]
            if fmt:
                c.number_format = fmt
            if col >= "R" or len(col) == 2:
                c.font = Font(color="595959")
        if MONTHS[r - FIRST_ROW].month == 12:
            for col in "ABCDEFGHIJKLMNOP":
                ws[f"{col}{r}"].border = Border(bottom=THIN)


def metric_formulas(ws, r):
    ws[f"K{r}"] = f"=IFERROR(F{r}/G{r},0)"
    ws[f"L{r}"] = f"=IFERROR(H{r}/G{r},0)"
    ws[f"M{r}"] = f"=IFERROR(H{r}/F{r},0)"
    ws[f"N{r}"] = f"=IFERROR(H{r}/E{r},0)"
    ws[f"O{r}"] = f"=IFERROR(I{r}/E{r},0)"
    ws[f"P{r}"] = f"=IFERROR(J{r}/E{r},0)"


def fy_block(ws):
    header(ws, 5, 1, "FY")
    header(ws, 5, 4, "Status")
    for col in "EFGHIJKLMNOP":
        header(ws, 5, ws[f"{col}1"].column, SITE_COLS[col][0])
    ws.row_dimensions[5].height = 45
    for i, y in enumerate(YEARS):
        r = 6 + i
        ws[f"A{r}"] = y
        ws[f"A{r}"].font = BOLD
        ws[f"D{r}"] = (f'=IF(DATE(A{r},12,1)<=LastActual,"Actual",'
                       f'IF(DATE(A{r},1,1)>LastActual,"Forecast","Blended"))')
        for col in "EFGHIJ":
            ws[f"{col}{r}"] = f"=SUMIFS({RNG(col)},{RNG('B')},$A{r})"
            ws[f"{col}{r}"].number_format = SITE_COLS[col][1]
        metric_formulas(ws, r)
        for col in "KLMNOP":
            ws[f"{col}{r}"].number_format = SITE_COLS[col][1]
        for col in "ADEFGHIJKLMNOP":
            ws[f"{col}{r}"].fill = LIGHT


def build_site(code, typ, state, host):
    ws = wb.create_sheet(code)
    ws["A1"] = f"{code}  |  {typ}  |  {state}" + (f"  |  sales from {host} (excluded from national sales)" if host else "")
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"], ws["B2"], ws["C2"], ws["D2"] = "Code", code, "State", state
    ws["B2"].font = ws["D2"].font = BOLD
    # run-rate levels
    ws["AF2"], ws["AF3"] = "Run-rate level:", "(last 6 mths / season idx)"
    ws["AF2"].font = BOLD
    ws["AF3"].font = HELP
    for col, act, si in (("AG", "W", "AD"), ("AH", "X", "AE"), ("AI", "Y", "AF")):
        ws[f"{col}2"] = (f"=AVERAGEIFS({RNG(act)},{ACT_WIN('RunRateMonths')})"
                         f"/AVERAGEIFS({RNG(si)},{ACT_WIN('RunRateMonths')})")
        ws[f"{col}2"].number_format = SITE_COLS[col][1]
        ws[f"{col}2"].font = BOLD
    fy_block(ws)
    for i, m in enumerate(MONTHS):
        r = FIRST_ROW + i
        act = f'$D{r}="Actual"'
        ws[f"A{r}"] = m
        ws[f"B{r}"] = f"=YEAR(A{r})"
        ws[f"C{r}"] = f"=MONTH(A{r})"
        ws[f"D{r}"] = f'=IF(A{r}<=LastActual,"Actual","Forecast")'
        if typ == "IDC":
            ws[f"E{r}"] = f"='{host}'!E{r}"
            ws[f"R{r}"] = f'=IF({act},E{r},"")'
        else:
            ws[f"E{r}"] = f"=IF({act},R{r},AB{r}*(1+SalesAdj))"
            ws[f"R{r}"] = f'=IF({act},{data_sum("C", r)},"")'
        ws[f"F{r}"] = f"=IF({act},S{r},E{r}*AG{r}*(1+CasesAdj))"
        ws[f"G{r}"] = f"=IF({act},T{r},IFERROR(H{r}/AH{r},0))"
        ws[f"H{r}"] = f"=IF({act},U{r},J{r}-I{r})"
        ws[f"I{r}"] = f"=IF({act},V{r},E{r}*AI{r})"
        ws[f"J{r}"] = f"=IF({act},U{r}+V{r},E{r}*AC{r})"
        metric_formulas(ws, r)
        for col, src in (("S", "D"), ("T", "E"), ("U", "F"), ("V", "G")):
            ws[f"{col}{r}"] = f'=IF({act},{data_sum(src, r)},"")'
        ws[f"W{r}"] = f'=IF({act},IFERROR(S{r}/R{r},""),"")'
        ws[f"X{r}"] = f'=IF({act},IFERROR(U{r}/T{r}/Z{r}/AA{r},""),"")'
        ws[f"Y{r}"] = f'=IF({act},IFERROR(V{r}/R{r},""),"")'
        ws[f"Z{r}"] = (f"=INDEX(Calendar!$H$5:$L${4 + N},MATCH($A{r},Calendar!$A$5:$A${4 + N},0),"
                       f"MATCH($D$2,Calendar!$H$4:$L$4,0))")
        ws[f"AA{r}"] = f"=INDEX(Calendar!$M$5:$M${4 + N},MATCH($A{r},Calendar!$A$5:$A${4 + N},0))"
        ws[f"AB{r}"] = f"=SUMIFS({PR('C')},{PR('A')},$A{r},{PR('B')},$B$2)"
        ws[f"AC{r}"] = f"=SUMIFS({PR('D')},{PR('A')},$A{r},{PR('B')},$B$2)"
        for col, src in (("AD", "W"), ("AE", "X"), ("AF", "Y")):
            ws[f"{col}{r}"] = (f"=IFERROR(AVERAGEIFS({RNG(src)},{ACT_WIN('SeasonMonths')},{RNG('C')},$C{r})"
                               f"/AVERAGEIFS({RNG(src)},{ACT_WIN('SeasonMonths')}),\"\")")
        ws[f"AG{r}"] = f"=IFERROR($AG$2*AD{r},\"\")"
        ws[f"AH{r}"] = f"=IFERROR($AH$2*AE{r}*Z{r}*AA{r},\"\")"
        ws[f"AI{r}"] = f"=IFERROR($AI$2*AF{r},\"\")"
    style_site_sheet(ws)


# national first so it sits after Checks; built after the site tabs exist
for s in SITES:
    build_site(*s)

ws = wb.create_sheet("National", index=wb.sheetnames.index("NSW-1"))
ws["A1"] = "National  |  8 regions + 2 IDCs  |  sales exclude IDCs"
ws["A1"].font = Font(bold=True, size=14)
ws["A2"] = "Model status"
ws["B2"] = "=Checks!B3"
ws["B2"].font = BOLD
fy_block(ws)
for i, m in enumerate(MONTHS):
    r = FIRST_ROW + i
    ws[f"A{r}"] = m
    ws[f"B{r}"] = f"=YEAR(A{r})"
    ws[f"C{r}"] = f"=MONTH(A{r})"
    ws[f"D{r}"] = f'=IF(A{r}<=LastActual,"Actual","Forecast")'
    ws[f"E{r}"] = "=" + "+".join(f"'{c}'!E{r}" for c in REGIONS)
    for col in "FGHIJ":
        ws[f"{col}{r}"] = "=" + "+".join(f"'{s[0]}'!{col}{r}" for s in SITES)
    metric_formulas(ws, r)
national_cols = {k: v for k, v in SITE_COLS.items() if k <= "P" and len(k) == 1}
for col, (title, fmt, width) in national_cols.items():
    ws.column_dimensions[col].width = width
    header(ws, 13, ws[f"{col}1"].column, title)
    if col in UNITS:
        u = ws[f"{col}14"]
        u.value, u.font, u.alignment, u.fill = UNITS[col], Font(italic=True), CENTER, LIGHT
ws.row_dimensions[13].height = 45
ws.merge_cells("E12:J12")
header(ws, 12, 5, "Drivers")
ws.merge_cells("K12:P12")
header(ws, 12, 11, "Metrics")
ws.freeze_panes = f"E{FIRST_ROW}"
ws.conditional_formatting.add(f"A{FIRST_ROW}:P{LAST_ROW}", FormulaRule(formula=[f'$D{FIRST_ROW}="Forecast"'], fill=LIGHT))
for r in range(FIRST_ROW, LAST_ROW + 1):
    for col, (_, fmt, _) in national_cols.items():
        if fmt:
            ws[f"{col}{r}"].number_format = fmt
    if MONTHS[r - FIRST_ROW].month == 12:
        for col in "ABCDEFGHIJKLMNOP":
            ws[f"{col}{r}"].border = Border(bottom=THIN)

# site x FY summary tables to the right
ws.column_dimensions["Q"].width = 3
ws.column_dimensions["R"].width = 10
tables = [
    ("Total payroll ($m)", lambda c, y: f"SUMIFS('{c}'!$J${FIRST_ROW}:$J${LAST_ROW},'{c}'!$B${FIRST_ROW}:$B${LAST_ROW},{y})", F_M),
    ("Payroll % of sales (IDC: % of host sales)",
     lambda c, y: (f"IFERROR(SUMIFS('{c}'!$J${FIRST_ROW}:$J${LAST_ROW},'{c}'!$B${FIRST_ROW}:$B${LAST_ROW},{y})"
                   f"/SUMIFS('{c}'!$E${FIRST_ROW}:$E${LAST_ROW},'{c}'!$B${FIRST_ROW}:$B${LAST_ROW},{y}),0)"), F_PCT),
    ("CPLH (cases / hr)",
     lambda c, y: (f"IFERROR(SUMIFS('{c}'!$F${FIRST_ROW}:$F${LAST_ROW},'{c}'!$B${FIRST_ROW}:$B${LAST_ROW},{y})"
                   f"/SUMIFS('{c}'!$G${FIRST_ROW}:$G${LAST_ROW},'{c}'!$B${FIRST_ROW}:$B${LAST_ROW},{y}),0)"), F_1),
    ("Cost per case ($, staff only)",
     lambda c, y: (f"IFERROR(SUMIFS('{c}'!$H${FIRST_ROW}:$H${LAST_ROW},'{c}'!$B${FIRST_ROW}:$B${LAST_ROW},{y})"
                   f"/SUMIFS('{c}'!$F${FIRST_ROW}:$F${LAST_ROW},'{c}'!$B${FIRST_ROW}:$B${LAST_ROW},{y}),0)"), F_CPC),
]
row = 4
for title, fn, fmt in tables:
    ws.cell(row=row, column=18, value=title).font = BOLD
    header(ws, row + 1, 18, "Site")
    for k, y in enumerate(YEARS):
        header(ws, row + 1, 19 + k, f"FY{y}", width=11)
    for i, (code, *_rest) in enumerate(SITES):
        ws.cell(row=row + 2 + i, column=18, value=code)
        for k, y in enumerate(YEARS):
            c = ws.cell(row=row + 2 + i, column=19 + k, value="=" + fn(code, y))
            c.number_format = fmt
    row += len(SITES) + 4

# ------------------------------------------------------------------ Checks
ws = wb.create_sheet("Checks", index=wb.sheetnames.index("National"))
ws.column_dimensions["A"].width = 62
for col in "BCD":
    ws.column_dimensions[col].width = 16
ws["A1"] = "Checks"
ws["A1"].font = Font(bold=True, size=14)
ws["A3"] = "Overall"
ws["B3"] = '=IF(COUNTIF(D6:D40,"ERROR")=0,"ALL CHECKS OK","CHECK ERRORS")'
ws["B3"].font = BOLD
for j, t in enumerate(["Check", "Value", "Expected", "Result"], 1):
    header(ws, 5, j, t)
checks = []
act_months = '(YEAR(LastActual)-2025)*12+MONTH(LastActual)'
fc_months = '(2028-YEAR(LastActual))*12+12-MONTH(LastActual)'
for code, *_r in SITES:
    checks.append((f"{code}: actual rows in Data_Actuals (one per month)",
                   f'=COUNTIFS({DR("B")},"{code}",{DR("A")},"<="&LastActual)', f"={act_months}", "eq"))
for code, *_r in SITES:
    checks.append((f"{code}: plan payroll % present for every forecast month",
                   f'=COUNTIFS({PR("B")},"{code}",{PR("A")},">"&LastActual,{PR("D")},">0")',
                   f"={fc_months}", "eq"))
checks.append(("Unknown site codes in Data_Actuals",
               f"=SUMPRODUCT(({DR('B')}<>\"\")*(COUNTIF(SiteCodes,{DR('B')})=0))", "=0", "eq"))
checks.append(("National sales = sum of 8 regions (IDCs excluded)",
               f"=SUM(National!E{FIRST_ROW}:E{LAST_ROW})",
               "=" + "+".join(f"SUM('{c}'!E{FIRST_ROW}:E{LAST_ROW})" for c in REGIONS), "eq"))
checks.append(("National total payroll = sum of all 10 sites",
               f"=SUM(National!J{FIRST_ROW}:J{LAST_ROW})",
               "=" + "+".join(f"SUM('{s[0]}'!J{FIRST_ROW}:J{LAST_ROW})" for s in SITES), "eq"))
checks.append(("Staff + leadership = total payroll (national, all months)",
               f"=SUM(National!H{FIRST_ROW}:I{LAST_ROW})", f"=SUM(National!J{FIRST_ROW}:J{LAST_ROW})", "eq"))
checks.append(("Months with negative staff payroll (leadership > total)",
               "=" + "+".join(f"COUNTIF('{s[0]}'!H{FIRST_ROW}:H{LAST_ROW},\"<0\")" for s in SITES), "=0", "eq"))
checks.append(("At least 12 actual months for seasonality", f"={act_months}", "=SeasonMonths", "ge"))
for i, (label, val, exp, op) in enumerate(checks):
    r = 6 + i
    ws[f"A{r}"], ws[f"B{r}"], ws[f"C{r}"] = label, val, exp
    test = f"ABS(B{r}-C{r})<=MAX(1,ABS(C{r})*0.000001)" if op == "eq" else f"B{r}>=C{r}"
    ws[f"D{r}"] = f'=IF({test},"OK","ERROR")'
    for col in "BC":
        ws[f"{col}{r}"].number_format = F_INT
ws.conditional_formatting.add("D6:D40", FormulaRule(formula=['D6="ERROR"'], fill=PatternFill("solid", fgColor="F8CBAD")))
ws.conditional_formatting.add("D6:D40", FormulaRule(formula=['D6="OK"'], fill=PatternFill("solid", fgColor="C6EFCE")))

# tab colours
for nm in wb.sheetnames:
    if nm in ("Inputs", "Data_Actuals", "Plan_Inputs", "Holidays"):
        wb[nm].sheet_properties.tabColor = "FFC000"
    elif nm in REGIONS:
        wb[nm].sheet_properties.tabColor = "1F3864"
    elif nm in ("IDC-NSW", "IDC-VIC"):
        wb[nm].sheet_properties.tabColor = "2E75B6"
    elif nm == "National":
        wb[nm].sheet_properties.tabColor = "00B050"

wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print("saved", OUT)

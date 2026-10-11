"""Build UK_House_Prices.xlsx: 70 years of Nationwide house prices, in cash and in real terms.

Every number in the workbook is a live Excel formula over the two raw sheets, so changing the base year, a
mortgage rate or the deposit recalculates everything. Raw data is pasted as published, untouched.

    python tools/build_workbook.py            # writes UK_House_Prices.xlsx (formulas, no cached values)
    python tools/check_workbook.py            # recalculates in LibreOffice and checks the numbers against pandas

Formulas stick to functions that Excel 2010+ and LibreOffice both evaluate (INDEX/MATCH, SUMIFS, AVERAGEIFS,
MAXIFS/MINIFS, PMT), so the file opens anywhere and can be verified outside Excel.
"""
import csv
import os
from datetime import datetime

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "UK_House_Prices.xlsx")
RAW_HP = os.path.join(ROOT, "data", "raw", "nationwide_uk_house_prices.csv")
RAW_CPI = os.path.join(ROOT, "data", "raw", "worldbank_cpi.csv")

# ---------------------------------------------------------------- style
FONT = "Arial"
INK, MUTED, ACCENT, WARN = "13201B", "6B7570", "0F766E", "C2410C"
f_base = Font(name=FONT, size=10, color=INK)
f_bold = Font(name=FONT, size=10, bold=True, color=INK)
f_head = Font(name=FONT, size=10, bold=True, color="FFFFFF")
f_title = Font(name=FONT, size=18, bold=True, color=INK)
f_sub = Font(name=FONT, size=10, color=MUTED)
f_input = Font(name=FONT, size=10, color="0000FF", bold=True)        # blue = something you can change
f_link = Font(name=FONT, size=10, color="008000")                    # green = pulled from another sheet
f_kpi = Font(name=FONT, size=20, bold=True, color=INK)
fill_head = PatternFill("solid", fgColor=INK)
fill_input = PatternFill("solid", fgColor="FFF2CC")
fill_band = PatternFill("solid", fgColor="F3F5F4")
thin = Side(style="thin", color="D5DAD7")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
GBP, GBP2, PCT, PCT0, MULT = '£#,##0;(£#,##0);"-"', '£#,##0.00;(£#,##0.00);"-"', '0.0%;-0.0%;"-"', '0%;-0%;"-"', '0.0"x"'


def style_range(ws, ref, font=None, fill=None, fmt=None, align=None, border=None):
    for row in ws[ref]:
        for c in row:
            if font:
                c.font = font
            if fill:
                c.fill = fill
            if fmt:
                c.number_format = fmt
            if align:
                c.alignment = align
            if border:
                c.border = border


def header(ws, row, labels, col=1):
    for i, h in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=h)
        c.font, c.fill = f_head, fill_head
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 30


def widths(ws, w):
    for i, x in enumerate(w, start=1):
        ws.column_dimensions[get_column_letter(i)].width = x


def name(wb, nm, ref):
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)


# ---------------------------------------------------------------- data
hp = list(csv.reader(open(RAW_HP, encoding="utf-8")))
cpi = list(csv.reader(open(RAW_CPI, encoding="utf-8")))
N_HP = len(hp) - 1
N_CPI = len(cpi) - 1
gbr_years = sorted(int(r[2]) for r in cpi[1:] if r[1] == "GBR")
FIRST_CPI, LAST_CPI = gbr_years[0], gbr_years[-1]
hp_years = sorted({int(r[0][:4]) for r in hp[1:]})

wb = Workbook()

# ================================================================ About
ab = wb.active
ab.title = "About"
ab.sheet_view.showGridLines = False
widths(ab, [3, 26, 100])
ab["B2"] = "Are UK houses really more expensive?"
ab["B2"].font = f_title
ab["B3"] = "Nationwide average house prices since 1953, in cash and after inflation. Every number is a live formula over the raw data."
ab["B3"].font = f_sub
rows = [
    ("How to use", ""),
    ("Dashboard", "Pick a base year (yellow cell). Every real price in the workbook is restated in that year's pounds."),
    ("Mortgage", "Compare the monthly cost of the average house in two quarters. Change the deposit, term and interest rates (yellow cells)."),
    ("Colour code", "Blue text in a yellow cell: an input you can change. Black: a formula. Green: a value pulled from another sheet."),
    ("", ""),
    ("Sheet map", ""),
    ("Dashboard", "Headline figures, charts and the biggest real-terms falls."),
    ("Mortgage", "Monthly repayment then vs now, and a rate x deposit sensitivity grid (PMT)."),
    ("Annual", "One row per year: average price in cash and in base-year pounds, inflation, growth."),
    ("Decades", "Growth per decade, in cash and real terms, with compound annual rates."),
    ("Quarterly", "One row per Nationwide quarter: real price, year-on-year change, running real peak and fall from peak."),
    ("CPI", "UK inflation each year from the World Bank, chained into a price index and rebased to the base year."),
    ("Checks", "Reconciliation tests. The Dashboard shows whether they all pass."),
    ("Raw_Nationwide", "Nationwide house price series exactly as published (datasets/house-prices-uk)."),
    ("Raw_CPI", "World Bank consumer price data for every country, exactly as published (datasets/cpi)."),
    ("", ""),
    ("Sources", ""),
    ("House prices", "Nationwide Building Society, UK house prices since 1952 (quarterly average price, all houses and new houses), via github.com/datasets/house-prices-uk (ODC PDDL 1.0)."),
    ("Inflation", "World Bank, consumer prices for the United Kingdom, via github.com/datasets/cpi (ODC PDDL 1.0)."),
    ("", ""),
    ("Notes and data quality", ""),
    ("CPI values", "The CPI file's metadata describes an index with 2005 = 100, but the values are annual % changes (UK 1975 = 24.2, 2024 = 3.3). The workbook treats them as annual inflation and chains them into an index (1959 = 100)."),
    ("Annual deflator", "Inflation is annual, so every quarter in a year is deflated by that year's average price level."),
    ("Coverage", f"Real prices run from {FIRST_CPI} to {LAST_CPI}, the years with inflation data. Quarters after {LAST_CPI} show cash prices only."),
    ("Methodology", "Nationwide's method changed over time: simple averages to 1959, floor-area weighting 1960-73, a regression-based mix adjustment from 1983. Long-run comparisons are approximate."),
    ("Quarter labels", "Nationwide dates its quarters in February, May, August and November; these are labelled Q1 to Q4."),
    ("Not included", "Earnings, so this is not an affordability ratio; and actual mortgage rates, which you enter yourself on the Mortgage sheet."),
]
r = 5
for a, b in rows:
    ab.cell(row=r, column=2, value=a).font = f_bold if b == "" else f_base
    if b == "" and a:
        ab.cell(row=r, column=2).font = Font(name=FONT, size=12, bold=True, color=ACCENT)
    c = ab.cell(row=r, column=3, value=b)
    c.font, c.alignment = f_base, Alignment(wrap_text=True, vertical="top")
    ab.cell(row=r, column=2).alignment = Alignment(vertical="top")
    r += 1
ab["B" + str(r + 1)] = "Built by Qais Adeniji-Bello · github.com/QaisBello/uk-house-prices-excel"
ab["B" + str(r + 1)].font = f_sub

# ================================================================ Raw sheets (as published)
raw_hp = wb.create_sheet("Raw_Nationwide")
for i, row in enumerate(hp):
    for j, v in enumerate(row):
        if i == 0:
            raw_hp.cell(row=1, column=j + 1, value=v)
        elif j == 0:
            raw_hp.cell(row=i + 1, column=1, value=datetime.strptime(v, "%Y-%m-%d")).number_format = "yyyy-mm-dd"
        else:
            raw_hp.cell(row=i + 1, column=j + 1, value=float(v) if v != "" else None)
header(raw_hp, 1, hp[0])
raw_hp.freeze_panes = "A2"
widths(raw_hp, [12] + [13] * 8)

raw_cpi = wb.create_sheet("Raw_CPI")
for i, row in enumerate(cpi):
    for j, v in enumerate(row):
        if i == 0 or j < 2:
            raw_cpi.cell(row=i + 1, column=j + 1, value=v)
        else:
            raw_cpi.cell(row=i + 1, column=j + 1, value=float(v) if j == 3 else int(v))
header(raw_cpi, 1, cpi[0])
raw_cpi.freeze_panes = "A2"
widths(raw_cpi, [34, 13, 8, 12])
for ws in (raw_hp, raw_cpi):
    ws.sheet_properties.tabColor = "B6BEB9"

# ================================================================ CPI (calc)
cs = wb.create_sheet("CPI", 1)
cs.sheet_view.showGridLines = False
cs["A1"] = "UK price level, chained from the World Bank's annual inflation"
cs["A1"].font = Font(name=FONT, size=13, bold=True, color=INK)
cs["A2"] = "Index: 1959 = 100, then each year = previous x (1 + inflation). Rebased: base year = 100 (base year set on the Dashboard)."
cs["A2"].font = f_sub
header(cs, 4, ["Year", "Inflation (World Bank, %)", "Price index (1959 = 100)", "Index (base year = 100)", "Rows found"])
c0 = 5
cs.cell(row=c0, column=1, value=FIRST_CPI - 1)
cs.cell(row=c0, column=2, value=None)
cs.cell(row=c0, column=3, value=100)
cs.cell(row=c0, column=3).comment = Comment("Starting point of the chained index. Any value works: only ratios are used.", "QAB")
for k, y in enumerate(range(FIRST_CPI, LAST_CPI + 1), start=1):
    rr = c0 + k
    cs.cell(row=rr, column=1, value=y)
    cs.cell(row=rr, column=2, value=f'=SUMIFS(Raw_CPI!$D:$D,Raw_CPI!$B:$B,"GBR",Raw_CPI!$C:$C,A{rr})/100')
    cs.cell(row=rr, column=3, value=f"=C{rr - 1}*(1+B{rr})")
    cs.cell(row=rr, column=5, value=f'=COUNTIFS(Raw_CPI!$B:$B,"GBR",Raw_CPI!$C:$C,A{rr})')
CPI_END = c0 + (LAST_CPI - FIRST_CPI + 1)   # row of the last year
for rr in range(c0, CPI_END + 1):
    cs.cell(row=rr, column=4, value=f"=C{rr}/INDEX($C${c0}:$C${CPI_END},MATCH(BaseYear,$A${c0}:$A${CPI_END},0))*100")
style_range(cs, f"A{c0}:E{CPI_END}", font=f_base, border=box)
style_range(cs, f"B{c0}:B{CPI_END}", fmt=PCT)
style_range(cs, f"C{c0}:D{CPI_END}", fmt="0.0")
style_range(cs, f"B{c0 + 1}:B{CPI_END}", font=f_link)
widths(cs, [10, 16, 16, 16, 10])
cs.freeze_panes = "A5"
cs.add_table(Table(displayName="tblCPI", ref=f"A4:E{CPI_END}", tableStyleInfo=TableStyleInfo(name="TableStyleLight1", showRowStripes=True)))
name(wb, "CPI_Year", f"CPI!$A${c0 + 1}:$A${CPI_END}")   # 1960 on: the 1959 anchor row has no inflation data
name(wb, "CPI_Index", f"CPI!$D${c0 + 1}:$D${CPI_END}")

# ================================================================ Quarterly (calc)
q = wb.create_sheet("Quarterly", 1)
q.sheet_view.showGridLines = False
q["A1"] = "Nationwide quarters: cash price, real price and fall from the real peak"
q["A1"].font = Font(name=FONT, size=13, bold=True, color=INK)
q["A2"] = "Real price = cash price restated in base-year pounds. Fall from peak = real price against the highest real price reached so far."
q["A2"].font = f_sub
cols = ["Date", "Year", "Quarter", "Label", "Average price, all houses (£)", "Average price, new houses (£)", "Price index (base = 100)",
        "Real price (base-year £)", "Change on a year earlier", "Real change on a year earlier", "Nationwide's published change",
        "Our change minus published (pp)", "Highest real price so far (£)", "Real fall from peak"]
header(q, 4, cols)
Q0 = 5
QN = Q0 + N_HP - 1
for k in range(N_HP):
    rr, src = Q0 + k, k + 2
    q.cell(row=rr, column=1, value=f"=Raw_Nationwide!A{src}")
    q.cell(row=rr, column=2, value=f"=YEAR(A{rr})")
    q.cell(row=rr, column=3, value=f'="Q"&ROUNDUP(MONTH(A{rr})/3,0)')
    q.cell(row=rr, column=4, value=f'=B{rr}&" "&C{rr}')
    q.cell(row=rr, column=5, value=f"=Raw_Nationwide!B{src}")
    q.cell(row=rr, column=6, value=f"=Raw_Nationwide!D{src}")
    q.cell(row=rr, column=7, value=f'=IFERROR(INDEX(CPI_Index,MATCH(B{rr},CPI_Year,0)),"")')
    q.cell(row=rr, column=8, value=f'=IF(G{rr}="","",E{rr}*100/G{rr})')
    if k >= 4:
        q.cell(row=rr, column=9, value=f"=E{rr}/E{rr - 4}-1")
        q.cell(row=rr, column=10, value=f'=IF(OR(H{rr}="",H{rr - 4}=""),"",H{rr}/H{rr - 4}-1)')
        q.cell(row=rr, column=12, value=f'=IF(K{rr}="","",(I{rr}-K{rr})*100)')
    q.cell(row=rr, column=11, value=f'=IF(Raw_Nationwide!C{src}="","",Raw_Nationwide!C{src}/100)')
    q.cell(row=rr, column=13, value=f'=IF(H{rr}="","",MAX(H${Q0}:H{rr}))')
    q.cell(row=rr, column=14, value=f'=IF(H{rr}="","",H{rr}/M{rr}-1)')
style_range(q, f"A{Q0}:N{QN}", font=f_base)
style_range(q, f"A{Q0}:A{QN}", fmt="mmm yyyy", font=f_link)
style_range(q, f"E{Q0}:F{QN}", fmt=GBP, font=f_link)
style_range(q, f"K{Q0}:K{QN}", font=f_link, fmt=PCT)
style_range(q, f"G{Q0}:G{QN}", fmt="0.0")
style_range(q, f"H{Q0}:H{QN}", fmt=GBP)
style_range(q, f"I{Q0}:J{QN}", fmt=PCT)
style_range(q, f"L{Q0}:L{QN}", fmt='0.00;-0.00;"-"')
style_range(q, f"M{Q0}:M{QN}", fmt=GBP)
style_range(q, f"N{Q0}:N{QN}", fmt=PCT)
widths(q, [11, 7, 8, 10, 14, 14, 12, 15, 12, 12, 13, 13, 15, 12])
q.freeze_panes = "E5"
q.add_table(Table(displayName="tblQuarterly", ref=f"A4:N{QN}", tableStyleInfo=TableStyleInfo(name="TableStyleLight1", showRowStripes=True)))
q.conditional_formatting.add(f"N{Q0}:N{QN}", CellIsRule(operator="lessThan", formula=["-0.2"], font=Font(name=FONT, color=WARN, bold=True)))
q.conditional_formatting.add(f"J{Q0}:J{QN}", ColorScaleRule(start_type="num", start_value=-0.2, start_color="F4B183",
                                                            mid_type="num", mid_value=0, mid_color="FFFFFF", end_type="num", end_value=0.2, end_color="9FD8CF"))
for nm, col in (("Q_Date", "A"), ("Q_Year", "B"), ("Q_Label", "D"), ("Q_Price", "E"), ("Q_Real", "H"), ("Q_Fall", "N"), ("Q_Diff", "L")):
    name(wb, nm, f"Quarterly!${col}${Q0}:${col}${QN}")
first_real_row = Q0 + next(i for i, r_ in enumerate(hp[1:]) if int(r_[0][:4]) == FIRST_CPI)
last_real_row = Q0 + max(i for i, r_ in enumerate(hp[1:]) if int(r_[0][:4]) == LAST_CPI)

# ================================================================ Annual (calc)
an = wb.create_sheet("Annual", 1)
an.sheet_view.showGridLines = False
an["A1"] = "One row per year (average of the year's quarters)"
an["A1"].font = Font(name=FONT, size=13, bold=True, color=INK)
an["A2"] = "Years with inflation data only, so cash and real prices can be compared."
an["A2"].font = f_sub
header(an, 4, ["Year", "Average price (£)", "Real price (base-year £)", "Inflation", "Cash growth", "Real growth", "Real price vs 1960"])
A0 = 5
years = list(range(FIRST_CPI, LAST_CPI + 1))
AN = A0 + len(years) - 1
for k, y in enumerate(years):
    rr = A0 + k
    an.cell(row=rr, column=1, value=y)
    an.cell(row=rr, column=2, value=f"=AVERAGEIFS(Q_Price,Q_Year,A{rr})")
    an.cell(row=rr, column=3, value=f"=AVERAGEIFS(Q_Real,Q_Year,A{rr})")
    an.cell(row=rr, column=4, value=f"=INDEX(CPI!$B:$B,MATCH(A{rr},CPI!$A:$A,0))")
    if k:
        an.cell(row=rr, column=5, value=f"=B{rr}/B{rr - 1}-1")
        an.cell(row=rr, column=6, value=f"=C{rr}/C{rr - 1}-1")
    an.cell(row=rr, column=7, value=f"=C{rr}/$C${A0}")
style_range(an, f"A{A0}:G{AN}", font=f_base, border=box)
style_range(an, f"B{A0}:C{AN}", fmt=GBP)
style_range(an, f"D{A0}:F{AN}", fmt=PCT)
style_range(an, f"D{A0}:D{AN}", font=f_link)
style_range(an, f"G{A0}:G{AN}", fmt=MULT)
widths(an, [8, 14, 16, 11, 11, 11, 13])
an.freeze_panes = "B5"
an.conditional_formatting.add(f"F{A0 + 1}:F{AN}", ColorScaleRule(start_type="num", start_value=-0.15, start_color="F4B183",
                                                                 mid_type="num", mid_value=0, mid_color="FFFFFF", end_type="num", end_value=0.15, end_color="9FD8CF"))
an.add_table(Table(displayName="tblAnnual", ref=f"A4:G{AN}", tableStyleInfo=TableStyleInfo(name="TableStyleLight1", showRowStripes=True)))
name(wb, "A_Year", f"Annual!$A${A0}:$A${AN}")
name(wb, "A_Price", f"Annual!$B${A0}:$B${AN}")
name(wb, "A_Real", f"Annual!$C${A0}:$C${AN}")

# ================================================================ Decades
de = wb.create_sheet("Decades", 1)
de.sheet_view.showGridLines = False
de["A1"] = "Growth by decade"
de["A1"].font = Font(name=FONT, size=13, bold=True, color=INK)
de["A2"] = f"First year of the decade against its last year (the 2020s run to {LAST_CPI}). Compound annual rate = (end / start)^(1 / years) - 1."
de["A2"].font = f_sub
header(de, 4, ["Decade", "From", "To", "Cash price at start (£)", "Cash price at end (£)", "Cash growth", "Real growth",
               "Cash growth a year", "Real growth a year", "Average inflation a year"])
D0 = 5
decs = list(range(1960, LAST_CPI + 1, 10))
DN = D0 + len(decs) - 1
for k, d in enumerate(decs):
    rr = D0 + k
    de.cell(row=rr, column=1, value=f"{d}s")
    de.cell(row=rr, column=2, value=d)
    de.cell(row=rr, column=3, value=f"=MIN(B{rr}+9,MAX(A_Year))")
    de.cell(row=rr, column=4, value=f"=INDEX(A_Price,MATCH(B{rr},A_Year,0))")
    de.cell(row=rr, column=5, value=f"=INDEX(A_Price,MATCH(C{rr},A_Year,0))")
    de.cell(row=rr, column=6, value=f"=E{rr}/D{rr}-1")
    de.cell(row=rr, column=7, value=f"=INDEX(A_Real,MATCH(C{rr},A_Year,0))/INDEX(A_Real,MATCH(B{rr},A_Year,0))-1")
    de.cell(row=rr, column=8, value=f"=(1+F{rr})^(1/(C{rr}-B{rr}))-1")
    de.cell(row=rr, column=9, value=f"=(1+G{rr})^(1/(C{rr}-B{rr}))-1")
    de.cell(row=rr, column=10, value=f"=(INDEX(CPI!$C:$C,MATCH(C{rr},CPI!$A:$A,0))/INDEX(CPI!$C:$C,MATCH(B{rr},CPI!$A:$A,0)))^(1/(C{rr}-B{rr}))-1")
style_range(de, f"A{D0}:J{DN}", font=f_base, border=box)
style_range(de, f"D{D0}:E{DN}", fmt=GBP)
style_range(de, f"F{D0}:J{DN}", fmt=PCT)
widths(de, [9, 7, 7, 15, 15, 12, 12, 12, 12, 13])
de.conditional_formatting.add(f"G{D0}:G{DN}", DataBarRule(start_type="num", start_value=-0.5, end_type="num", end_value=1.5, color="0F766E"))
de.freeze_panes = "B5"

# ================================================================ Checks
ck = wb.create_sheet("Checks")
ck.sheet_view.showGridLines = False
ck["A1"] = "Reconciliation checks"
ck["A1"].font = Font(name=FONT, size=13, bold=True, color=INK)
ck["A2"] = "Each test compares the calculation sheets with the raw data. PASS or FAIL updates with every recalculation."
ck["A2"].font = f_sub
header(ck, 4, ["Check", "Result", "Status", "What it proves"])
checks = [
    ("Quarters in the calculation = rows in the raw file", f"=COUNT(Q_Price)&\" of \"&COUNT(Raw_Nationwide!B:B)", f"=IF(COUNT(Q_Price)=COUNT(Raw_Nationwide!B:B),\"PASS\",\"FAIL\")",
     "No quarter dropped or duplicated."),
    ("Sum of cash prices = sum in the raw file", "=SUM(Q_Price)-SUM(Raw_Nationwide!B:B)", "=IF(ABS(B6)<0.5,\"PASS\",\"FAIL\")", "The calculation reads every price unchanged."),
    ("Largest gap between our year-on-year change and Nationwide's (pp)", "=MAX(MAX(Q_Diff),-MIN(Q_Diff))", "=IF(B7<0.1,\"PASS\",\"FAIL\")",
     "Our change formula reproduces Nationwide's published figure to within rounding (0.1 percentage points)."),
    ("UK inflation years found once each in the raw CPI file", f"=SUM(CPI!E{c0 + 1}:E{CPI_END})&\" of {LAST_CPI - FIRST_CPI + 1}\"",
     f"=IF(AND(SUM(CPI!E{c0 + 1}:E{CPI_END})={LAST_CPI - FIRST_CPI + 1},MAX(CPI!E{c0 + 1}:E{CPI_END})=1),\"PASS\",\"FAIL\")",
     "Every year from 1960 matches exactly one World Bank row for the UK."),
    ("Price index in the base year = 100", "=INDEX(CPI_Index,MATCH(BaseYear,CPI_Year,0))", "=IF(ABS(B9-100)<0.000001,\"PASS\",\"FAIL\")", "Real prices are in base-year pounds."),
    ("Real price = cash price in the base year", "=INDEX(A_Real,MATCH(BaseYear,A_Year,0))-INDEX(A_Price,MATCH(BaseYear,A_Year,0))",
     "=IF(ABS(B10)<0.5,\"PASS\",\"FAIL\")", "Deflating is consistent: in the base year, nothing changes."),
    ("Quarters with a real price", f"=COUNT(Q_Real)", f"=IF(COUNT(Q_Real)=COUNTIFS(Q_Year,\">={FIRST_CPI}\",Q_Year,\"<={LAST_CPI}\"),\"PASS\",\"FAIL\")",
     f"Every quarter from {FIRST_CPI} to {LAST_CPI} has a real price; later quarters are left blank on purpose."),
    ("Mortgage payment check", "=Mortgage!C14*Mortgage!C9*12-Mortgage!C12", "=IF(B12>0,\"PASS\",\"FAIL\")",
     "Total repaid over the term is more than the loan, so the PMT signs are right."),
]
for k, (a, b, c, d) in enumerate(checks):
    rr = 5 + k
    ck.cell(row=rr, column=1, value=a)
    ck.cell(row=rr, column=2, value=b)
    ck.cell(row=rr, column=3, value=c)
    ck.cell(row=rr, column=4, value=d)
CKN = 5 + len(checks) - 1
style_range(ck, f"A5:D{CKN}", font=f_base, border=box, align=Alignment(wrap_text=True, vertical="top"))
style_range(ck, f"B5:B{CKN}", fmt="General")
ck["B9"].number_format = "0.0"
ck.conditional_formatting.add(f"C5:C{CKN}", CellIsRule(operator="equal", formula=['"PASS"'], font=Font(name=FONT, bold=True, color="16845F")))
ck.conditional_formatting.add(f"C5:C{CKN}", CellIsRule(operator="equal", formula=['"FAIL"'], font=Font(name=FONT, bold=True, color="FFFFFF"), fill=PatternFill("solid", fgColor=WARN)))
widths(ck, [52, 14, 9, 70])

# ================================================================ Mortgage
mo = wb.create_sheet("Mortgage", 1)
mo.sheet_view.showGridLines = False
widths(mo, [3, 38, 16, 16, 3, 16, 12, 12, 12, 12, 12])
mo["B2"] = "What did the average house cost a month?"
mo["B2"].font = f_title
mo["B3"] = "Repayment mortgage on the average Nationwide house in two quarters. Yellow cells are yours to change."
mo["B3"].font = f_sub
labels_ok = [f"{int(r_[0][:4])} Q{(int(r_[0][5:7]) + 2) // 3}" for r_ in hp[1:] if int(r_[0][:4]) >= FIRST_CPI]
default_a, default_b = "2007 Q3", [l for l in labels_ok if l.startswith(str(LAST_CPI))][-1]   # latest quarter with inflation data
grid = [
    (5, "", "Then", "Now"),
    (6, "Quarter", default_a, default_b),
    (7, "Mortgage rate (your assumption)", 0.05, 0.05),
    (8, "Deposit", 0.10, 0.10),
    (9, "Term (years)", 25, 25),
]
for rr, lab, a, b in grid:
    mo.cell(row=rr, column=2, value=lab)
    mo.cell(row=rr, column=3, value=a)
    mo.cell(row=rr, column=4, value=b)
style_range(mo, "B5:D5", font=f_head, fill=fill_head, align=Alignment(horizontal="center"))
style_range(mo, "C6:D9", font=f_input, fill=fill_input, border=box, align=Alignment(horizontal="center"))
style_range(mo, "C7:D8", fmt=PCT)
mo["C7"].comment = Comment("An assumption for illustration, not a historical rate. Enter the rate you want to test; both default to 5% so the comparison isolates the price.", "QAB")
calc = [
    (10, "Average house price (£)", '=INDEX(Q_Price,MATCH({c}6,Q_Label,0))', GBP),
    (11, "Deposit (£)", "={c}10*{c}8", GBP),
    (12, "Loan (£)", "={c}10-{c}11", GBP),
    (13, "Monthly rate", "={c}7/12", '0.000%'),
    (14, "Monthly repayment (£)", "=PMT({c}13,{c}9*12,-{c}12)", GBP2),
    (15, "Total repaid over the term (£)", "={c}14*{c}9*12", GBP),
    (16, "Monthly repayment in base-year £", '=IFERROR({c}14*100/INDEX(CPI_Index,MATCH(VALUE(LEFT({c}6,4)),CPI_Year,0)),"no inflation data yet")', GBP2),
    (17, "Deposit in base-year £", '=IFERROR({c}11*100/INDEX(CPI_Index,MATCH(VALUE(LEFT({c}6,4)),CPI_Year,0)),"no inflation data yet")', GBP),
]
for rr, lab, f, fmt in calc:
    mo.cell(row=rr, column=2, value=lab)
    for c in "CD":
        cell = mo[f"{c}{rr}"]
        cell.value = f.format(c=c)
        cell.number_format, cell.font, cell.border = fmt, f_base, box
        cell.alignment = Alignment(horizontal="right")
for r_ in range(6, 23):
    if mo[f"B{r_}"].value:
        mo[f"B{r_}"].font = f_base
mo["B14"].font = f_bold
style_range(mo, "C14:D14", font=f_bold)
mo["B19"] = "Change, then to now"
mo["B20"] = "Monthly repayment, cash"
mo["C20"] = "=D14/C14-1"
mo["B21"] = "Monthly repayment, real (base-year £)"
mo["C21"] = '=IFERROR(D16/C16-1,"needs inflation data for both quarters")'
mo["B22"] = "Deposit, real (base-year £)"
mo["C22"] = '=IFERROR(D17/C17-1,"needs inflation data for both quarters")'
style_range(mo, "B20:C22", font=f_base, border=box)
style_range(mo, "C20:C22", fmt=PCT)
mo["B19"].font = f_bold
dv_q = DataValidation(type="list", formula1=f"=Quarterly!$D${first_real_row}:$D${QN}", allow_blank=False)
dv_q.error, dv_q.errorTitle = "Pick a quarter from the list, such as 2007 Q3.", "Quarter"
mo.add_data_validation(dv_q)
dv_q.add("C6:D6")
dv_r = DataValidation(type="decimal", operator="between", formula1="0.0001", formula2="0.25")
dv_r.error = "Enter a rate between 0.01% and 25%."
mo.add_data_validation(dv_r)
dv_r.add("C7:D7")
dv_d = DataValidation(type="decimal", operator="between", formula1="0", formula2="0.95")
dv_d.error = "Enter a deposit between 0% and 95%."
mo.add_data_validation(dv_d)
dv_d.add("C8:D8")
dv_t = DataValidation(type="whole", operator="between", formula1="5", formula2="40")
mo.add_data_validation(dv_t)
dv_t.add("C9:D9")

# sensitivity: monthly repayment on the 'Now' house, rate (rows) x deposit (columns)
mo["F5"] = "Monthly repayment on the 'Now' house: rate x deposit"
mo["F5"].font = f_bold
mo["F6"] = "Rate \\ deposit"
mo["F6"].font, mo["F6"].fill = f_head, fill_head
deps = [0.05, 0.10, 0.15, 0.20, 0.25]
rates = [0.02, 0.03, 0.04, 0.05, 0.06, 0.07, 0.08]
for j, dpt in enumerate(deps):
    c = mo.cell(row=6, column=7 + j, value=dpt)
    c.font, c.fill, c.number_format, c.alignment = f_head, fill_head, PCT0, Alignment(horizontal="center")
for i, rt in enumerate(rates):
    rr = 7 + i
    c = mo.cell(row=rr, column=6, value=rt)
    c.font, c.fill, c.number_format, c.alignment, c.border = f_input, fill_input, PCT, Alignment(horizontal="center"), box
    for j in range(len(deps)):
        col = get_column_letter(7 + j)
        cell = mo.cell(row=rr, column=7 + j, value=f"=PMT($F{rr}/12,$D$9*12,-$D$10*(1-{col}$6))")
        cell.number_format, cell.font, cell.border = GBP, f_base, box
mo["F15"] = "Term from cell D9. Rates in column F are editable."
mo["F15"].font = f_sub
mo.conditional_formatting.add("G7:K13", ColorScaleRule(start_type="min", start_color="E6F2F0", end_type="max", end_color="F4B183"))

# ================================================================ Dashboard
db = wb.create_sheet("Dashboard", 1)
db.sheet_view.showGridLines = False
widths(db, [3, 30, 18, 3, 26, 18, 3, 26, 18, 3, 14])
db["B2"] = "Are UK houses really more expensive?"
db["B2"].font = f_title
db["B3"] = f"Nationwide average house price, {hp_years[0]}-{hp_years[-1]}, in cash and in base-year pounds (after inflation)."
db["B3"].font = f_sub
db["B5"] = "Base year (pick a year)"
db["B5"].font = f_bold
db["C5"] = LAST_CPI
db["C5"].font, db["C5"].fill, db["C5"].border, db["C5"].alignment = f_input, fill_input, box, Alignment(horizontal="center")
db["C5"].comment = Comment("Every real price in the workbook is restated in this year's pounds.", "QAB")
name(wb, "BaseYear", "Dashboard!$C$5")
dv_y = DataValidation(type="list", formula1=f"=CPI!$A${c0 + 1}:$A${CPI_END}", allow_blank=False)
dv_y.error = f"Pick a year from {FIRST_CPI} to {LAST_CPI}."
db.add_data_validation(dv_y)
dv_y.add("C5")
db["E5"] = "Checks"
db["E5"].font = f_bold
db["F5"] = '=IF(COUNTIF(Checks!C:C,"FAIL")=0,"All "&COUNTIF(Checks!C:C,"PASS")&" pass","FAILURES: see Checks")'
db["F5"].font = Font(name=FONT, size=10, bold=True, color="16845F")
db.conditional_formatting.add("F5", FormulaRule(formula=['COUNTIF(Checks!C:C,"FAIL")>0'], font=Font(name=FONT, bold=True, color=WARN)))

kpis = [
    ("B", 7, "Latest average price", "=INDEX(Q_Price,MATCH(9.99E+307,Q_Price))", GBP, '="Nationwide, "&INDEX(Q_Label,MATCH(9.99E+307,Q_Price))'),
    ("E", 7, "Cash price, 1960 to latest", "=INDEX(Q_Price,MATCH(9.99E+307,Q_Price))/INDEX(A_Price,1)", MULT, '="from "&TEXT(INDEX(A_Price,1),"£#,##0")&" (1960 average)"'),
    ("H", 7, f"Real price, 1960 to {LAST_CPI}", f"=INDEX(A_Real,MATCH({LAST_CPI},A_Year,0))/INDEX(A_Real,1)", MULT,
     '="in "&BaseYear&" pounds: "&TEXT(INDEX(A_Real,1),"£#,##0")&" to "&TEXT(INDEX(A_Real,ROWS(A_Real)),"£#,##0")'),
    ("B", 11, "Real growth a year", f"=(INDEX(A_Real,ROWS(A_Real))/INDEX(A_Real,1))^(1/({LAST_CPI}-{FIRST_CPI}))-1", PCT, f'="compound, {FIRST_CPI}-{LAST_CPI}, above inflation"'),
    ("E", 11, "Inflation a year", f"=(INDEX(CPI!$C:$C,MATCH({LAST_CPI},CPI!$A:$A,0))/INDEX(CPI!$C:$C,MATCH({FIRST_CPI},CPI!$A:$A,0)))^(1/({LAST_CPI}-{FIRST_CPI}))-1", PCT,
     '="World Bank, UK, same period"'),
    ("H", 11, "Biggest real fall from a peak", "=MIN(Q_Fall)", PCT, '="trough "&INDEX(Q_Label,MATCH(MIN(Q_Fall),Q_Fall,0))'),
]
for col, rr, lab, f, fmt, sub in kpis:
    nxt = get_column_letter(ord(col) - 64 + 1)
    db[f"{col}{rr}"] = lab
    db[f"{col}{rr}"].font = Font(name=FONT, size=10, bold=True, color=MUTED)
    db[f"{col}{rr + 1}"] = f
    db[f"{col}{rr + 1}"].font, db[f"{col}{rr + 1}"].number_format = f_kpi, fmt
    db[f"{col}{rr + 2}"] = sub
    db[f"{col}{rr + 2}"].font = f_sub
    for r_ in range(rr, rr + 3):
        for cc in (col, nxt):
            db[f"{cc}{r_}"].fill = fill_band
    db.merge_cells(f"{col}{rr + 1}:{nxt}{rr + 1}")
    db.merge_cells(f"{col}{rr + 2}:{nxt}{rr + 2}")
    db.row_dimensions[rr + 1].height = 30

# falls table
db["B15"] = "The big real-terms falls"
db["B15"].font = Font(name=FONT, size=12, bold=True, color=ACCENT)
header(db, 16, ["Episode", "Real fall from peak", "", "Trough"], col=2)
db.merge_cells("C16:D16")
episodes = [("1970s oil shock", 1973, 1979), ("Early-1990s crash", 1989, 1996), ("Financial crisis", 2007, 2013)]
for k, (lab, a, b) in enumerate(episodes):
    rr = 17 + k
    db[f"B{rr}"] = f"{lab} ({a}-{b})"
    db[f"C{rr}"] = f'=_xlfn.MINIFS(Q_Fall,Q_Year,">={a}",Q_Year,"<={b}")'
    db[f"E{rr}"] = f"=INDEX(Q_Label,MATCH(C{rr},Q_Fall,0))"
    db.merge_cells(f"C{rr}:D{rr}")
    for c in "BCDE":
        db[f"{c}{rr}"].font, db[f"{c}{rr}"].border = f_base, box
    db[f"C{rr}"].number_format = PCT
    db[f"C{rr}"].font = Font(name=FONT, size=10, bold=True, color=WARN)
db["B20"] = "Fall from peak: real price against the highest real price reached before it. Cash prices fell far less in each episode, because inflation did part of the work."
db["B20"].font = f_sub

# charts
lc = LineChart()
lc.title = "Average price: cash and in base-year pounds"
lc.y_axis.title = "£"
lc.y_axis.numFmt = "£#,##0"
lc.height, lc.width = 8.5, 17
lc.add_data(Reference(an, min_col=2, max_col=3, min_row=4, max_row=AN), titles_from_data=True)
lc.set_categories(Reference(an, min_col=1, min_row=A0, max_row=AN))
lc.series[0].graphicalProperties.line.solidFill = "B6BEB9"
lc.series[0].graphicalProperties.line.width = 22000
lc.series[1].graphicalProperties.line.solidFill = ACCENT
lc.series[1].graphicalProperties.line.width = 28000
lc.x_axis.tickLblSkip = 5
lc.legend.position = "b"
lc.x_axis.delete = False
lc.y_axis.delete = False
for s_ in lc.series:
    s_.smooth = False
db.add_chart(lc, "B22")

bc = BarChart()
bc.type = "col"
bc.title = "Real growth by decade"
bc.y_axis.numFmt = "0%"
bc.height, bc.width = 8.5, 12
bc.add_data(Reference(de, min_col=7, min_row=4, max_row=DN), titles_from_data=True)
bc.set_categories(Reference(de, min_col=1, min_row=D0, max_row=DN))
bc.series[0].graphicalProperties = GraphicalProperties(solidFill=ACCENT)
bc.legend = None
bc.x_axis.delete = False
bc.y_axis.delete = False
bc.x_axis.tickLblPos = "low"
db.add_chart(bc, "H22")

dc = LineChart()
dc.title = "Real price against its previous peak"
dc.y_axis.numFmt = "0%"
dc.height, dc.width = 8.5, 29.5
dc.add_data(Reference(q, min_col=14, min_row=first_real_row, max_row=last_real_row), titles_from_data=False)
dc.set_categories(Reference(q, min_col=4, min_row=first_real_row, max_row=last_real_row))
dc.series[0].graphicalProperties.line.solidFill = WARN
dc.series[0].graphicalProperties.line.width = 22000
dc.legend = None
dc.x_axis.tickLblSkip = 20
dc.x_axis.delete = False
dc.y_axis.delete = False
dc.series[0].smooth = False
dc.x_axis.tickLblPos = "low"
dc.x_axis.tickLblSkip = 40
db.add_chart(dc, "B40")

# order of tabs
order = ["About", "Dashboard", "Mortgage", "Annual", "Decades", "Quarterly", "CPI", "Checks", "Raw_Nationwide", "Raw_CPI"]
wb._sheets = [wb[s] for s in order]
for s, color in (("Dashboard", ACCENT), ("Mortgage", ACCENT), ("Checks", "16845F")):
    wb[s].sheet_properties.tabColor = color
wb.active = 1
for ws in wb.worksheets:
    ws.sheet_view.zoomScale = 100
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print("wrote", OUT, f"({N_HP} quarters, UK inflation {FIRST_CPI}-{LAST_CPI})")

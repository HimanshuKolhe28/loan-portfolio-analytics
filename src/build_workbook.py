"""
Loan Portfolio Risk Analytics — Excel workbook generator.

Generates a synthetic retail/SME loan book (seeded, reproducible) and builds a
formula-driven Excel workbook with:
  * RBI IRACP-style asset classification (SMA-0/1/2, Sub-standard, Doubtful, Loss)
  * Provisioning, Gross/Net NPA, Provision Coverage Ratio
  * Interactive dashboard (Region / Product dropdown filters) with charts
  * Top-10 NPA exposure leaderboard (LARGE + INDEX/MATCH)
  * Vintage (cohort) analysis with heat-map formatting
  * EMI calculator with full amortization schedule (PMT / IPMT / PPMT)

Every number in the workbook is an Excel formula driven by the Loans and
Assumptions sheets — change an input and the whole model recalculates.

Usage:
    python src/build_workbook.py            # writes output/Loan_Portfolio_Risk_Analytics.xlsx
    python src/build_workbook.py --rows 2000 --seed 7
"""
from __future__ import annotations

import argparse
import random
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

# --------------------------------------------------------------------------- #
# Styling
# --------------------------------------------------------------------------- #
FONT = "Arial"
NAVY = "1F3864"
TEAL = "2E75B6"
LIGHT = "DDEBF7"
GREY = "F2F2F2"

F_TITLE = Font(name=FONT, size=18, bold=True, color=NAVY)
F_SUB = Font(name=FONT, size=10, italic=True, color="595959")
F_HDR = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_BODY = Font(name=FONT, size=10)
F_BOLD = Font(name=FONT, size=10, bold=True)
F_INPUT = Font(name=FONT, size=10, color="0000FF")          # hardcoded input
F_LINK = Font(name=FONT, size=10, color="008000")           # cross-sheet link
F_KPI = Font(name=FONT, size=16, bold=True, color=NAVY)
F_KPI_LBL = Font(name=FONT, size=9, bold=True, color="595959")

FILL_HDR = PatternFill("solid", fgColor=NAVY)
FILL_SEC = PatternFill("solid", fgColor=TEAL)
FILL_LIGHT = PatternFill("solid", fgColor=LIGHT)
FILL_GREY = PatternFill("solid", fgColor=GREY)
FILL_INPUT = PatternFill("solid", fgColor="FFFF00")

THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center")

INR = '₹#,##0;(₹#,##0);"-"'
CR = '₹#,##0.00" Cr";(₹#,##0.00" Cr");"-"'
PCT = '0.00%;(0.00%);"-"'
INT = '#,##0;(#,##0);"-"'

AS_OF = date(2026, 6, 30)

# --------------------------------------------------------------------------- #
# Synthetic data
# --------------------------------------------------------------------------- #
BRANCHES = [
    ("BLR-01", "Bengaluru", "South"), ("CHN-01", "Chennai", "South"),
    ("HYD-01", "Hyderabad", "South"), ("KOC-01", "Kochi", "South"),
    ("MUM-01", "Mumbai", "West"), ("PUN-01", "Pune", "West"),
    ("AMD-01", "Ahmedabad", "West"), ("DEL-01", "New Delhi", "North"),
    ("LKO-01", "Lucknow", "North"), ("JAI-01", "Jaipur", "North"),
    ("KOL-01", "Kolkata", "East"), ("BBS-01", "Bhubaneswar", "East"),
    ("PAT-01", "Patna", "East"), ("GUW-01", "Guwahati", "East"),
]

# product: (min_amt, max_amt, min_rate, max_rate, tenures, weight)
PRODUCTS = {
    "Home Loan":     (2_000_000, 15_000_000, 0.084, 0.098, [120, 180, 240], 0.22),
    "Auto Loan":     (400_000, 2_000_000, 0.089, 0.120, [36, 48, 60, 84], 0.20),
    "Personal Loan": (100_000, 1_500_000, 0.105, 0.180, [12, 24, 36, 48, 60], 0.28),
    "Business Loan": (500_000, 5_000_000, 0.130, 0.200, [12, 24, 36, 60], 0.18),
    "Gold Loan":     (50_000, 500_000, 0.090, 0.120, [6, 12], 0.12),
}

FIRST = ["Aarav", "Vivaan", "Aditya", "Arjun", "Sai", "Reyansh", "Ishaan", "Kabir",
         "Ananya", "Diya", "Priya", "Sneha", "Kavya", "Meera", "Riya", "Pooja",
         "Rahul", "Rohan", "Karthik", "Suresh", "Lakshmi", "Deepa", "Farhan", "Zoya",
         "Manish", "Nikhil", "Tanvi", "Gaurav", "Swati", "Harsh"]
LAST = ["Sharma", "Verma", "Iyer", "Nair", "Reddy", "Patel", "Das", "Mohanty",
        "Gupta", "Singh", "Rao", "Menon", "Banerjee", "Khan", "Joshi", "Pillai",
        "Chatterjee", "Kulkarni", "Mishra", "Sahoo"]


def generate_loans(n: int, seed: int) -> list[dict]:
    rng = random.Random(seed)
    names, weights = list(PRODUCTS), [p[5] for p in PRODUCTS.values()]
    loans = []
    for i in range(1, n + 1):
        prod = rng.choices(names, weights)[0]
        lo, hi, rlo, rhi, tenures, _ = PRODUCTS[prod]
        tenure = rng.choice(tenures)
        # disbursed within the last 36 months, but not after as-of date
        disb = AS_OF - timedelta(days=rng.randint(30, 1095))
        cibil = int(min(900, max(300, rng.gauss(735, 55))))
        # delinquency probability rises as CIBIL falls and for unsecured products
        base = {"Personal Loan": 0.13, "Business Loan": 0.15, "Auto Loan": 0.07,
                "Home Loan": 0.035, "Gold Loan": 0.05}[prod]
        p_delinq = base * (1 + max(0, 720 - cibil) / 60)
        if rng.random() < p_delinq:
            dpd = rng.choice([rng.randint(1, 30), rng.randint(31, 60), rng.randint(61, 90),
                              rng.randint(91, 365), rng.randint(91, 365),
                              rng.randint(366, 730), rng.randint(731, 900)])
        else:
            dpd = 0
        # keep DPD consistent with loan age: a loan can't be overdue longer than it
        # has existed, and a fully matured loan is treated as closed (DPD 0)
        elapsed = (AS_OF.year - disb.year) * 12 + AS_OF.month - disb.month
        if elapsed >= tenure:
            dpd = 0
        dpd = min(dpd, max(0, elapsed * 30 - 15))
        loans.append({
            "id": f"LN{100000 + i}",
            "name": f"{rng.choice(FIRST)} {rng.choice(LAST)}",
            "branch": rng.choice(BRANCHES)[0],
            "product": prod,
            "disb": disb,
            "amount": round(rng.uniform(lo, hi), -3),
            "rate": round(rng.uniform(rlo, rhi), 4),
            "tenure": tenure,
            "cibil": cibil,
            "dpd": dpd,
        })
    return loans


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def header_row(ws, row, col, labels, fill=FILL_HDR, height=30):
    for j, lbl in enumerate(labels):
        c = ws.cell(row=row, column=col + j, value=lbl)
        c.font, c.fill, c.alignment, c.border = F_HDR, fill, CENTER, BOX
    ws.row_dimensions[row].height = height


def section(ws, ref, text):
    ws[ref] = text
    ws[ref].font = Font(name=FONT, size=12, bold=True, color=NAVY)


def widths(ws, mapping):
    for col, w in mapping.items():
        ws.column_dimensions[col].width = w


def body(cell, fmt=None, font=F_BODY, align=None):
    cell.font, cell.border = font, BOX
    if fmt:
        cell.number_format = fmt
    if align:
        cell.alignment = align


# --------------------------------------------------------------------------- #
# Sheets
# --------------------------------------------------------------------------- #
def build_readme(wb):
    ws = wb.active
    ws.title = "Guide"
    ws.sheet_view.showGridLines = False
    widths(ws, {"A": 3, "B": 28, "C": 95})
    ws["B2"] = "Loan Portfolio Risk Analytics"
    ws["B2"].font = F_TITLE
    ws["B3"] = "Formula-driven credit-risk model for a synthetic retail & SME loan book (India / BFSI)"
    ws["B3"].font = F_SUB

    rows = [
        ("SHEET", "WHAT IT DOES"),
        ("Dashboard", "Interactive KPIs & charts. Pick Region / Product in the yellow cells (C4, C5) — every figure, chart and the Top-10 table updates."),
        ("Loans", "Loan-level book. Blue columns are raw inputs; black columns are formulas (EMI, outstanding, DPD bucket, asset class, provision, risk grade)."),
        ("Vintage", "Cohort analysis by disbursal quarter: count, disbursed, GNPA % — heat-mapped to spot bad vintages."),
        ("EMI Calculator", "Standalone EMI tool with full amortization schedule (PMT / IPMT / PPMT) and principal-vs-interest chart."),
        ("Assumptions", "As-of date, RBI-style asset-classification & provisioning table, CIBIL risk-grade bands. Edit here to re-run the model."),
        ("Branches", "Branch master used to map each loan to City / Region via INDEX-MATCH."),
        ("", ""),
        ("COLOUR KEY", ""),
        ("Blue text", "Hardcoded input — safe to edit."),
        ("Black text", "Formula — do not overwrite."),
        ("Green text", "Link pulled from another sheet."),
        ("Yellow fill", "Key assumption / user selection."),
        ("", ""),
        ("TECHNIQUES USED", "SUMIFS / COUNTIFS / SUMPRODUCT, INDEX-MATCH (exact & approximate), LARGE with tie-break keys, PMT / PV / IPMT / PPMT, EDATE, "
                            "named ranges, data validation dropdowns, dynamic filter flag column, conditional formatting (data bars, colour scales, formula rules), "
                            "Excel Tables, combo & doughnut charts, frozen panes, cell comments documenting assumptions."),
        ("DISCLAIMER", "All data is synthetic (seeded random). Provisioning rates are simplified illustrations of RBI IRACP norms, not regulatory advice."),
    ]
    for i, (a, b) in enumerate(rows, start=5):
        ws.cell(row=i, column=2, value=a).font = F_BOLD
        ws.cell(row=i, column=3, value=b).font = F_BODY
        ws.cell(row=i, column=3).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=i, column=2).alignment = Alignment(vertical="top")
    header_row(ws, 5, 2, ["SHEET", "WHAT IT DOES"])
    ws["B14"].font = F_INPUT
    ws["B15"].font = F_BODY
    ws["B16"].font = F_LINK
    ws["B17"].fill = FILL_INPUT
    ws.row_dimensions[19].height = 45
    ws.row_dimensions[20].height = 30


def build_assumptions(wb):
    ws = wb.create_sheet("Assumptions")
    ws.sheet_view.showGridLines = False
    widths(ws, {"A": 3, "B": 24, "C": 18, "D": 16, "E": 14, "F": 12, "G": 3, "H": 16, "I": 14})
    ws["B2"] = "Model Assumptions"
    ws["B2"].font = F_TITLE

    ws["B4"] = "As-of (reporting) date"
    ws["B4"].font = F_BOLD
    ws["C4"] = AS_OF
    ws["C4"].number_format = "dd-mmm-yyyy"
    ws["C4"].font, ws["C4"].fill, ws["C4"].border = F_INPUT, FILL_INPUT, BOX
    ws["C4"].comment = Comment("Portfolio snapshot date. All ageing, outstanding and "
                               "vintage calculations are measured to this date.", "Model")
    wb.defined_names["AsOfDate"] = DefinedName("AsOfDate", attr_text="Assumptions!$C$4")

    section(ws, "B6", "Asset Classification & Provisioning (by DPD)")
    header_row(ws, 7, 2, ["DPD From", "Bucket", "Asset Class", "NPA?", "Provision %"])
    table = [
        (0, "Current", "Standard", "No", 0.004),
        (1, "SMA-0", "Standard", "No", 0.004),
        (31, "SMA-1", "Standard", "No", 0.004),
        (61, "SMA-2", "Standard", "No", 0.004),
        (91, "Sub-standard", "NPA", "Yes", 0.15),
        (366, "Doubtful", "NPA", "Yes", 0.40),
        (731, "Loss", "NPA", "Yes", 1.00),
    ]
    for i, r in enumerate(table, start=8):
        for j, v in enumerate(r):
            c = ws.cell(row=i, column=2 + j, value=v)
            body(c, PCT if j == 4 else None, F_INPUT, CENTER)
    ws["B16"] = ("Source: simplified from RBI Master Circular on IRACP norms (SMA classification & "
                 "NPA at 90+ DPD). Doubtful provision blended at 40%; illustrative only.")
    ws["B16"].font = F_SUB
    for name, rng in [("DPD_From", "$B$8:$B$14"), ("DPD_Bucket", "$C$8:$C$14"),
                      ("DPD_Class", "$D$8:$D$14"), ("DPD_NPA", "$E$8:$E$14"),
                      ("DPD_Prov", "$F$8:$F$14")]:
        wb.defined_names[name] = DefinedName(name, attr_text=f"Assumptions!{rng}")

    section(ws, "H6", "Risk Grade by CIBIL")
    header_row(ws, 7, 8, ["CIBIL From", "Grade"])
    for i, (s, g) in enumerate([(300, "D - High Risk"), (650, "C - Watch"),
                                (700, "B - Moderate"), (750, "A - Prime")], start=8):
        body(ws.cell(row=i, column=8, value=s), None, F_INPUT, CENTER)
        body(ws.cell(row=i, column=9, value=g), None, F_INPUT, CENTER)
    ws.column_dimensions["I"].width = 16
    wb.defined_names["CIBIL_From"] = DefinedName("CIBIL_From", attr_text="Assumptions!$H$8:$H$11")
    wb.defined_names["CIBIL_Grade"] = DefinedName("CIBIL_Grade", attr_text="Assumptions!$I$8:$I$11")
    ws["H13"] = "Source: model convention (bands chosen by analyst)."
    ws["H13"].font = F_SUB
    return ws


def build_branches(wb):
    ws = wb.create_sheet("Branches")
    widths(ws, {"A": 12, "B": 16, "C": 10})
    header_row(ws, 1, 1, ["Branch", "City", "Region"])
    for i, (b, c, r) in enumerate(BRANCHES, start=2):
        for j, v in enumerate((b, c, r), start=1):
            body(ws.cell(row=i, column=j, value=v), font=F_INPUT)
    ws.freeze_panes = "A2"
    return len(BRANCHES) + 1


LOAN_COLS = [
    # (header, width, kind)  kind: in = input, f = formula, l = lookup/link
    ("Loan ID", 11, "in"), ("Customer", 18, "in"), ("Branch", 9, "in"),
    ("Region", 9, "l"), ("Product", 14, "in"), ("Disbursal Date", 12, "in"),
    ("Loan Amount (₹)", 14, "in"), ("Rate (p.a.)", 9, "in"), ("Tenure (m)", 8, "in"),
    ("CIBIL", 7, "in"), ("DPD", 7, "in"),
    ("EMI (₹)", 12, "f"), ("Months Elapsed", 9, "f"), ("EMIs Paid", 8, "f"),
    ("Outstanding (₹)", 14, "f"), ("Overdue (₹)", 12, "f"), ("DPD Bucket", 12, "f"),
    ("Asset Class", 10, "f"), ("NPA?", 6, "f"), ("Provision %", 9, "f"),
    ("Provision (₹)", 12, "f"), ("Risk Grade", 13, "f"), ("Vintage Qtr", 10, "f"),
    ("In Filter", 7, "f"), ("NPA Rank Key", 12, "f"),
]


def build_loans(wb, loans, n_branch_rows):
    ws = wb.create_sheet("Loans")
    last = len(loans) + 1
    for j, (h, w, _) in enumerate(LOAN_COLS, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w
    header_row(ws, 1, 1, [c[0] for c in LOAN_COLS], height=34)

    br = f"Branches!$A$2:$A${n_branch_rows}"
    rg = f"Branches!$C$2:$C${n_branch_rows}"
    for i, ln in enumerate(loans, start=2):
        r = str(i)
        vals = [
            ln["id"], ln["name"], ln["branch"],
            f"=INDEX({rg},MATCH(C{r},{br},0))",
            ln["product"], ln["disb"], ln["amount"], ln["rate"], ln["tenure"],
            ln["cibil"], ln["dpd"],
            f"=ROUND(PMT(H{r}/12,I{r},-G{r}),0)",                                    # L EMI
            f"=MAX(0,(YEAR(AsOfDate)-YEAR(F{r}))*12+MONTH(AsOfDate)-MONTH(F{r}))",    # M elapsed
            f"=MAX(0,MIN(I{r},M{r})-INT(K{r}/30))",                                  # N paid
            f"=IF(N{r}>=I{r},0,ROUND(PV(H{r}/12,I{r}-N{r},-L{r}),0))",               # O outstanding
            f"=MIN(I{r},M{r})*L{r}-N{r}*L{r}",                                       # P overdue
            f"=INDEX(DPD_Bucket,MATCH(K{r},DPD_From,1))",                            # Q bucket
            f"=INDEX(DPD_Class,MATCH(K{r},DPD_From,1))",                             # R class
            f"=INDEX(DPD_NPA,MATCH(K{r},DPD_From,1))",                               # S NPA?
            f"=INDEX(DPD_Prov,MATCH(K{r},DPD_From,1))",                              # T prov %
            f"=ROUND(O{r}*T{r},0)",                                                  # U prov ₹
            f"=INDEX(CIBIL_Grade,MATCH(J{r},CIBIL_From,1))",                         # V grade
            f'=YEAR(F{r})&"-Q"&ROUNDUP(MONTH(F{r})/3,0)',                             # W vintage
            (f'=IF(AND(OR(Dashboard!$C$4="All",D{r}=Dashboard!$C$4),'
             f'OR(Dashboard!$C$5="All",E{r}=Dashboard!$C$5)),1,0)'),                  # X filter
            f'=IF(AND(X{r}=1,S{r}="Yes"),O{r}+ROW()/1000000,0)',                      # Y rank key
        ]
        for j, v in enumerate(vals, start=1):
            c = ws.cell(row=i, column=j, value=v)
            kind = LOAN_COLS[j - 1][2]
            c.font = F_INPUT if kind == "in" else (F_LINK if kind == "l" else F_BODY)
        ws[f"F{r}"].number_format = "dd-mmm-yy"
        for col in "GLOPU":
            ws[f"{col}{r}"].number_format = INR
        ws[f"H{r}"].number_format = "0.00%"
        ws[f"T{r}"].number_format = "0.0%"
        ws[f"Y{r}"].number_format = "#,##0"

    tab = Table(displayName="tblLoans", ref=f"A1:{get_column_letter(len(LOAN_COLS))}{last}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
    ws.add_table(tab)
    ws.freeze_panes = "C2"

    # conditional formatting on asset class
    red = PatternFill("solid", fgColor="F8CBAD")
    amber = PatternFill("solid", fgColor="FFE699")
    ws.conditional_formatting.add(
        f"Q2:Q{last}", FormulaRule(formula=[f'$S2="Yes"'], fill=red,
                                   font=Font(name=FONT, color="9C0006", bold=True)))
    ws.conditional_formatting.add(
        f"Q2:Q{last}", FormulaRule(formula=[f'AND($K2>0,$S2="No")'], fill=amber))
    ws.conditional_formatting.add(
        f"J2:J{last}", ColorScaleRule(start_type="num", start_value=550, start_color="F8696B",
                                      mid_type="num", mid_value=720, mid_color="FFEB84",
                                      end_type="num", end_value=850, end_color="63BE7B"))

    # data validation for inputs
    dv_prod = DataValidation(type="list", formula1='"' + ",".join(PRODUCTS) + '"', allow_blank=False)
    dv_br = DataValidation(type="list", formula1=f"={br}", allow_blank=False)
    dv_cibil = DataValidation(type="whole", operator="between", formula1="300", formula2="900",
                              error="CIBIL must be 300–900", showErrorMessage=True)
    dv_dpd = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="0",
                            error="DPD cannot be negative", showErrorMessage=True)
    for dv, rng in [(dv_prod, f"E2:E{last}"), (dv_br, f"C2:C{last}"),
                    (dv_cibil, f"J2:J{last}"), (dv_dpd, f"K2:K{last}")]:
        ws.add_data_validation(dv)
        dv.add(rng)

    ws["X1"].comment = Comment("1 when the loan matches the Dashboard Region/Product "
                               "selection. Drives every Dashboard KPI.", "Model")
    ws["Y1"].comment = Comment("Outstanding + ROW()/1e6 for filtered NPA loans — the tiny "
                               "row offset breaks ties so LARGE/MATCH returns unique rows.", "Model")
    ws["N1"].comment = Comment("Assumes each 30 DPD = one missed EMI.", "Model")
    return last


def build_dashboard(wb, last, loans):
    ws = wb.create_sheet("Dashboard", 1)
    ws.sheet_view.showGridLines = False
    widths(ws, {"A": 2, "B": 20, "C": 16, "D": 16, "E": 14, "F": 14, "G": 3,
                "H": 16, "I": 16, "J": 14, "K": 14, "L": 14, "M": 3})
    L = lambda col: f"Loans!${col}$2:${col}${last}"  # noqa: E731
    FLT = L("X")

    ws["B2"] = "Loan Portfolio Risk Dashboard"
    ws["B2"].font = F_TITLE
    ws["B3"] = '="Snapshot as of "&TEXT(AsOfDate,"dd-mmm-yyyy")&"  |  All amounts in ₹ Crore"'
    ws["B3"].font = F_SUB

    # Filters
    for ref, lbl, opts in [("4", "Region", "All,North,South,East,West"),
                           ("5", "Product", "All," + ",".join(PRODUCTS))]:
        ws[f"B{ref}"] = lbl
        ws[f"B{ref}"].font = F_BOLD
        c = ws[f"C{ref}"]
        c.value = "All"
        c.font, c.fill, c.border, c.alignment = Font(name=FONT, bold=True, color="0000FF"), FILL_INPUT, BOX, CENTER
        dv = DataValidation(type="list", formula1=f'"{opts}"', allow_blank=False)
        ws.add_data_validation(dv)
        dv.add(f"C{ref}")
    ws["D4"] = "◀ select filters"
    ws["D4"].font = F_SUB

    # KPI cards: label row 7, value row 8 ; second band rows 10/11
    kpis = [
        ("B", "ACTIVE LOANS", f'=COUNTIFS({FLT},1,{L("O")},">0")', INT),
        ("C", "DISBURSED", f"=SUMIFS({L('G')},{FLT},1)/10^7", CR),
        ("D", "OUTSTANDING (AUM)", f"=SUMIFS({L('O')},{FLT},1)/10^7", CR),
        ("E", "GROSS NPA", f'=SUMIFS({L("O")},{FLT},1,{L("S")},"Yes")/10^7', CR),
        ("F", "GNPA %", "=IFERROR(E8/D8,0)", PCT),
        ("H", "TOTAL PROVISION", f"=SUMIFS({L('U')},{FLT},1)/10^7", CR),
        ("I", "NET NPA", f'=E8-SUMIFS({L("U")},{FLT},1,{L("S")},"Yes")/10^7', CR),
        ("J", "NNPA %", '=IFERROR(I8/(D8-(E8-I8)),0)', PCT),
        ("K", "PROVISION COVERAGE", "=IFERROR((E8-I8)/E8,0)", PCT),
        ("L", "SMA (EARLY STRESS)", f'=SUMIFS({L("O")},{FLT},1,{L("K")},">0",{L("S")},"No")/10^7', CR),
    ]
    for col, lbl, f, fmt in kpis:
        a, b = ws[f"{col}7"], ws[f"{col}8"]
        a.value, b.value = lbl, f
        a.font, a.fill, a.alignment = F_KPI_LBL, FILL_LIGHT, CENTER
        b.font, b.fill, b.alignment, b.number_format = F_KPI, FILL_LIGHT, CENTER, fmt
        for c in (a, b):
            c.border = BOX
    ws.row_dimensions[7].height = 26
    ws.row_dimensions[8].height = 34

    second = [
        ("B", "WTD AVG RATE", f"=IFERROR(SUMPRODUCT({FLT},{L('O')},{L('H')})/SUMPRODUCT({FLT},{L('O')}),0)", PCT),
        ("C", "AVG CIBIL", f"=IFERROR(AVERAGEIFS({L('J')},{FLT},1),0)", "0"),
        ("D", "AVG TICKET (₹ Lakh)", f"=IFERROR(SUMIFS({L('G')},{FLT},1)/COUNTIFS({FLT},1)/10^5,0)", "#,##0.00"),
        ("E", "NPA ACCOUNTS", f'=COUNTIFS({FLT},1,{L("S")},"Yes")', INT),
        ("F", "OVERDUE AMT", f"=SUMIFS({L('P')},{FLT},1)/10^7", CR),
    ]
    for col, lbl, f, fmt in second:
        a, b = ws[f"{col}10"], ws[f"{col}11"]
        a.value, b.value = lbl, f
        a.font, a.fill, a.alignment = F_KPI_LBL, FILL_GREY, CENTER
        b.font, b.fill, b.alignment, b.number_format = Font(name=FONT, size=13, bold=True, color=NAVY), FILL_GREY, CENTER, fmt
        for c in (a, b):
            c.border = BOX
    ws.row_dimensions[10].height = 26
    ws.row_dimensions[11].height = 28
    ws["H10"] = "Traffic light: GNPA % > 5% turns red, 3–5% amber, < 3% green."
    ws["H10"].font = F_SUB
    red_f = PatternFill("solid", fgColor="F8CBAD")
    amb_f = PatternFill("solid", fgColor="FFE699")
    grn_f = PatternFill("solid", fgColor="C6EFCE")
    ws.conditional_formatting.add("F8", CellIsRule(operator="greaterThan", formula=["0.05"], fill=red_f))
    ws.conditional_formatting.add("F8", CellIsRule(operator="between", formula=["0.03", "0.05"], fill=amb_f))
    ws.conditional_formatting.add("F8", CellIsRule(operator="lessThan", formula=["0.03"], fill=grn_f))

    # --- DPD bucket table (B14) ---
    section(ws, "B13", "Book by DPD Bucket")
    header_row(ws, 14, 2, ["Bucket", "Accounts", "Outstanding", "% of Book", "Provision"])
    for i in range(7):
        r = 15 + i
        ws[f"B{r}"] = f"=INDEX(DPD_Bucket,{i + 1})"
        ws[f"C{r}"] = f"=COUNTIFS({FLT},1,{L('Q')},B{r},{L('O')},\">0\")"
        ws[f"D{r}"] = f"=SUMIFS({L('O')},{FLT},1,{L('Q')},B{r})/10^7"
        ws[f"E{r}"] = f"=IFERROR(D{r}/$D$8,0)"
        ws[f"F{r}"] = f"=SUMIFS({L('U')},{FLT},1,{L('Q')},B{r})/10^7"
        body(ws[f"B{r}"], font=F_LINK)
        body(ws[f"C{r}"], INT)
        body(ws[f"D{r}"], CR)
        body(ws[f"E{r}"], PCT)
        body(ws[f"F{r}"], CR)
    ws["B22"] = "Total"
    for col, fmt in [("C", INT), ("D", CR), ("E", PCT), ("F", CR)]:
        ws[f"{col}22"] = f"=SUM({col}15:{col}21)"
        body(ws[f"{col}22"], fmt, F_BOLD)
        ws[f"{col}22"].fill = FILL_LIGHT
    body(ws["B22"], font=F_BOLD)
    ws["B22"].fill = FILL_LIGHT
    ws["B23"] = '=IF(ABS(D22-D8)<0.0001,"✔ Reconciles to AUM","✖ Check: bucket total ≠ AUM")'
    ws["B23"].font = Font(name=FONT, size=9, italic=True, color="008000")
    ws.conditional_formatting.add("D15:D21", DataBarRule(start_type="num", start_value=0,
                                                         end_type="max", color="5B9BD5"))

    # --- Product table (H14) ---
    section(ws, "H13", "Performance by Product")
    header_row(ws, 14, 8, ["Product", "Outstanding", "GNPA", "GNPA %", "Wtd Rate"])
    for i, p in enumerate(PRODUCTS):
        r = 15 + i
        ws[f"H{r}"] = p
        ws[f"I{r}"] = f"=SUMIFS({L('O')},{FLT},1,{L('E')},H{r})/10^7"
        ws[f"J{r}"] = f'=SUMIFS({L("O")},{FLT},1,{L("E")},H{r},{L("S")},"Yes")/10^7'
        ws[f"K{r}"] = f"=IFERROR(J{r}/I{r},0)"
        ws[f"L{r}"] = (f'=IFERROR(SUMPRODUCT({FLT},--({L("E")}=H{r}),{L("O")},{L("H")})'
                       f'/SUMPRODUCT({FLT},--({L("E")}=H{r}),{L("O")}),0)')
        body(ws[f"H{r}"])
        body(ws[f"I{r}"], CR)
        body(ws[f"J{r}"], CR)
        body(ws[f"K{r}"], PCT)
        body(ws[f"L{r}"], PCT)
    ws.conditional_formatting.add("K15:K19", ColorScaleRule(start_type="min", start_color="63BE7B",
                                                            mid_type="percentile", mid_value=50, mid_color="FFEB84",
                                                            end_type="max", end_color="F8696B"))

    # --- Region table (H23) ---
    section(ws, "H22", "Performance by Region")
    header_row(ws, 23, 8, ["Region", "Outstanding", "GNPA", "GNPA %", "Accounts"])
    for i, reg in enumerate(["North", "South", "East", "West"]):
        r = 24 + i
        ws[f"H{r}"] = reg
        ws[f"I{r}"] = f"=SUMIFS({L('O')},{FLT},1,{L('D')},H{r})/10^7"
        ws[f"J{r}"] = f'=SUMIFS({L("O")},{FLT},1,{L("D")},H{r},{L("S")},"Yes")/10^7'
        ws[f"K{r}"] = f"=IFERROR(J{r}/I{r},0)"
        ws[f"L{r}"] = f"=COUNTIFS({FLT},1,{L('D')},H{r})"
        body(ws[f"H{r}"])
        body(ws[f"I{r}"], CR)
        body(ws[f"J{r}"], CR)
        body(ws[f"K{r}"], PCT)
        body(ws[f"L{r}"], INT)

    # --- Top 10 NPA exposures (B26) ---
    section(ws, "B29", "Top 10 NPA Exposures (current filter)")
    for k in range(1, 11):
        r = 30 + k
        key = f"LARGE({L('Y')},B{r})"
        m = f"MATCH({key},{L('Y')},0)"
        ws[f"B{r}"] = k
        ws[f"C{r}"] = f'=IF({key}=0,"",INDEX({L("A")},{m}))'
        ws[f"D{r}"] = f'=IF(C{r}="","",INDEX({L("B")},{m}))'
        ws[f"E{r}"] = f'=IF(C{r}="","",INDEX({L("E")},{m}))'
        ws[f"F{r}"] = f'=IF(C{r}="","",INDEX({L("O")},{m}))'
        ws[f"H{r}"] = f'=IF(C{r}="","",INDEX({L("K")},{m}))'
        ws[f"I{r}"] = f'=IF(C{r}="","",INDEX({L("Q")},{m}))'
        ws[f"J{r}"] = f'=IF(C{r}="","",INDEX({L("C")},{m}))'
        ws[f"K{r}"] = f'=IF(C{r}="","",INDEX({L("J")},{m}))'
        for col in "BCDEFHIJK":
            body(ws[f"{col}{r}"], align=CENTER if col in "BHK" else None)
        ws[f"F{r}"].number_format = INR
    header_row(ws, 30, 2, ["Rank", "Loan ID", "Customer", "Product", "Outstanding (₹)"])
    header_row(ws, 30, 8, ["DPD", "Asset Class", "Branch", "CIBIL"])
    ws.column_dimensions["G"].width = 3

    # --- Charts ---
    ch1 = BarChart()
    ch1.type = "col"
    ch1.title = "Outstanding by DPD Bucket (₹ Cr)"
    ch1.add_data(Reference(ws, min_col=4, min_row=14, max_row=21), titles_from_data=True)
    ch1.set_categories(Reference(ws, min_col=2, min_row=15, max_row=21))
    ch1.legend = None
    ch1.height, ch1.width = 7.5, 15
    ch1.y_axis.numFmt = "0"
    ch1.dataLabels = DataLabelList()
    ch1.dataLabels.showVal = True
    ch1.dataLabels.numFmt = "0.0"
    ws.add_chart(ch1, "N2")

    ch2 = BarChart()
    ch2.type = "bar"
    ch2.title = "GNPA % by Product"
    ch2.add_data(Reference(ws, min_col=11, min_row=14, max_row=19), titles_from_data=True)
    ch2.set_categories(Reference(ws, min_col=8, min_row=15, max_row=19))
    ch2.legend = None
    ch2.height, ch2.width = 7.5, 15
    ch2.x_axis.numFmt = "0%"
    ch2.dataLabels = DataLabelList()
    ch2.dataLabels.showVal = True
    ch2.dataLabels.numFmt = "0.0%"
    ws.add_chart(ch2, "N18")

    ch3 = DoughnutChart()
    ch3.title = "AUM Mix by Region"
    ch3.add_data(Reference(ws, min_col=9, min_row=23, max_row=27), titles_from_data=True)
    ch3.set_categories(Reference(ws, min_col=8, min_row=24, max_row=27))
    ch3.height, ch3.width = 7.5, 11
    ch3.dataLabels = DataLabelList()
    ch3.dataLabels.showPercent = True
    ws.add_chart(ch3, "N34")

    ws.freeze_panes = "A6"


def build_vintage(wb, last, loans):
    ws = wb.create_sheet("Vintage")
    ws.sheet_view.showGridLines = False
    widths(ws, {"A": 2, "B": 12, "C": 11, "D": 16, "E": 16, "F": 16, "G": 11, "H": 11})
    L = lambda col: f"Loans!${col}$2:${col}${last}"  # noqa: E731
    ws["B2"] = "Vintage (Cohort) Analysis"
    ws["B2"].font = F_TITLE
    ws["B3"] = "How each disbursal quarter is performing today — later-dated bad cohorts signal underwriting drift."
    ws["B3"].font = F_SUB
    header_row(ws, 5, 2, ["Vintage", "Loans", "Disbursed (₹ Cr)", "Outstanding (₹ Cr)",
                          "GNPA (₹ Cr)", "GNPA %", "Avg CIBIL"])
    qtrs = sorted({f"{l['disb'].year}-Q{(l['disb'].month - 1) // 3 + 1}" for l in loans})
    for i, q in enumerate(qtrs):
        r = 6 + i
        ws[f"B{r}"] = q
        ws[f"C{r}"] = f"=COUNTIFS({L('W')},B{r})"
        ws[f"D{r}"] = f"=SUMIFS({L('G')},{L('W')},B{r})/10^7"
        ws[f"E{r}"] = f"=SUMIFS({L('O')},{L('W')},B{r})/10^7"
        ws[f"F{r}"] = f'=SUMIFS({L("O")},{L("W")},B{r},{L("S")},"Yes")/10^7'
        ws[f"G{r}"] = f"=IFERROR(F{r}/E{r},0)"
        ws[f"H{r}"] = f"=IFERROR(AVERAGEIFS({L('J')},{L('W')},B{r}),0)"
        body(ws[f"B{r}"], align=CENTER)
        body(ws[f"C{r}"], INT)
        body(ws[f"D{r}"], "#,##0.00")
        body(ws[f"E{r}"], "#,##0.00")
        body(ws[f"F{r}"], "#,##0.00")
        body(ws[f"G{r}"], PCT)
        body(ws[f"H{r}"], "0")
    end = 5 + len(qtrs)
    tr = end + 1
    ws[f"B{tr}"] = "Total"
    for col in "CDEF":
        ws[f"{col}{tr}"] = f"=SUM({col}6:{col}{end})"
        body(ws[f"{col}{tr}"], INT if col == "C" else "#,##0.00", F_BOLD)
        ws[f"{col}{tr}"].fill = FILL_LIGHT
    ws[f"G{tr}"] = f"=IFERROR(F{tr}/E{tr},0)"
    body(ws[f"G{tr}"], PCT, F_BOLD)
    body(ws[f"B{tr}"], font=F_BOLD)
    ws[f"B{tr}"].fill = ws[f"G{tr}"].fill = FILL_LIGHT
    ws.conditional_formatting.add(f"G6:G{end}", ColorScaleRule(
        start_type="min", start_color="63BE7B", mid_type="percentile", mid_value=50,
        mid_color="FFEB84", end_type="max", end_color="F8696B"))
    ws.conditional_formatting.add(f"D6:D{end}", DataBarRule(start_type="num", start_value=0,
                                                            end_type="max", color="5B9BD5"))

    ch = BarChart()
    ch.type = "col"
    ch.title = "Disbursals vs GNPA % by Vintage"
    ch.add_data(Reference(ws, min_col=4, min_row=5, max_row=end), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=6, max_row=end))
    ch.y_axis.title = "₹ Cr"
    ch.y_axis.majorGridlines = None
    ln = LineChart()
    ln.add_data(Reference(ws, min_col=7, min_row=5, max_row=end), titles_from_data=True)
    ln.y_axis.axId = 200
    ln.y_axis.title = "GNPA %"
    ln.y_axis.numFmt = "0%"
    ln.y_axis.crosses = "max"
    ch += ln
    ch.height, ch.width = 9, 20
    ws.add_chart(ch, "J5")
    ws.freeze_panes = "A6"


def build_emi(wb):
    ws = wb.create_sheet("EMI Calculator")
    ws.sheet_view.showGridLines = False
    widths(ws, {"A": 2, "B": 22, "C": 16, "D": 3, "E": 8, "F": 12, "G": 15,
                "H": 13, "I": 13, "J": 13, "K": 15})
    ws["B2"] = "EMI Calculator & Amortization"
    ws["B2"].font = F_TITLE

    inputs = [("Loan Amount (₹)", 2_500_000, INR), ("Interest Rate (p.a.)", 0.0875, "0.00%"),
              ("Tenure (months)", 120, "0"), ("First EMI Date", date(2026, 7, 5), "dd-mmm-yyyy")]
    for i, (lbl, v, fmt) in enumerate(inputs, start=4):
        ws[f"B{i}"] = lbl
        ws[f"B{i}"].font = F_BOLD
        c = ws[f"C{i}"]
        c.value, c.number_format = v, fmt
        c.font, c.fill, c.border = F_INPUT, FILL_INPUT, BOX
    ws["C4"].comment = Comment("Example input — replace with any amount.", "Model")
    dv = DataValidation(type="whole", operator="between", formula1="1", formula2="360",
                        error="Tenure must be 1–360 months", showErrorMessage=True)
    ws.add_data_validation(dv)
    dv.add("C6")

    outs = [("Monthly EMI (₹)", "=IFERROR(PMT(C5/12,C6,-C4),0)", INR),
            ("Total Interest (₹)", "=C9*C6-C4", INR),
            ("Total Payable (₹)", "=C9*C6", INR),
            ("Interest / Principal", "=IFERROR(C10/C4,0)", "0.0%"),
            ("Last EMI Date", "=EDATE(C7,C6-1)", "dd-mmm-yyyy")]
    for i, (lbl, f, fmt) in enumerate(outs, start=9):
        ws[f"B{i}"] = lbl
        ws[f"B{i}"].font = F_BOLD
        c = ws[f"C{i}"]
        c.value, c.number_format = f, fmt
        c.font, c.border, c.fill = Font(name=FONT, bold=True, size=11, color=NAVY), BOX, FILL_LIGHT
    ws["B15"] = "Schedule shows up to 360 months; rows beyond tenure stay blank."
    ws["B15"].font = F_SUB

    header_row(ws, 3, 5, ["#", "Date", "Opening (₹)", "EMI (₹)", "Interest (₹)",
                          "Principal (₹)", "Closing (₹)"])
    for n in range(1, 361):
        r = 3 + n
        ws[f"E{r}"] = f'=IF({n}>$C$6,"",{n})'
        ws[f"F{r}"] = f'=IF(E{r}="","",EDATE($C$7,E{r}-1))'
        ws[f"G{r}"] = f'=IF(E{r}="","",{"$C$4" if n == 1 else f"K{r - 1}"})'
        ws[f"H{r}"] = f'=IF(E{r}="","",$C$9)'
        ws[f"I{r}"] = f'=IF(E{r}="","",-IPMT($C$5/12,E{r},$C$6,$C$4))'
        ws[f"J{r}"] = f'=IF(E{r}="","",-PPMT($C$5/12,E{r},$C$6,$C$4))'
        ws[f"K{r}"] = f'=IF(E{r}="","",MAX(0,G{r}-J{r}))'
        ws[f"F{r}"].number_format = "mmm-yy"
        for col in "GHIJK":
            ws[f"{col}{r}"].number_format = '₹#,##0'
        for col in "EFGHIJK":
            ws[f"{col}{r}"].font = F_BODY
    ws.freeze_panes = "E4"

    ch = LineChart()
    ch.title = "Interest vs Principal per EMI"
    ch.add_data(Reference(ws, min_col=9, max_col=10, min_row=3, max_row=123), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=5, min_row=4, max_row=123))
    ch.height, ch.width = 8, 15
    ch.y_axis.numFmt = "#,##0"
    ws.add_chart(ch, "B17")


# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--rows", type=int, default=1000)
    ap.add_argument("--seed", type=int, default=42)
    ap.add_argument("--out", default=str(Path(__file__).resolve().parents[1] / "output" /
                                         "Loan_Portfolio_Risk_Analytics.xlsx"))
    args = ap.parse_args()

    loans = generate_loans(args.rows, args.seed)
    wb = Workbook()
    build_readme(wb)
    build_assumptions(wb)
    n_br = build_branches(wb)
    last = build_loans(wb, loans, n_br)
    build_dashboard(wb, last, loans)
    build_vintage(wb, last, loans)
    build_emi(wb)

    order = ["Guide", "Dashboard", "Loans", "Vintage", "EMI Calculator", "Assumptions", "Branches"]
    wb._sheets = [wb[n] for n in order]
    wb.active = 1
    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    wb.save(args.out)
    print(f"Wrote {args.out} ({len(loans)} loans)")


if __name__ == "__main__":
    main()

# 📊 Loan Portfolio Risk Analytics — Advanced Excel Model

A fully formula-driven **credit-risk analytics workbook** for a retail & SME loan book, modelled on how Indian banks and NBFCs monitor asset quality under **RBI IRACP norms**. The workbook is generated reproducibly with Python (`openpyxl`), so the repo demonstrates both advanced Excel and engineering practice.

> **17,700+ live formulas · 0 errors · 1,000 loans · 7 sheets · 5 charts** — change any input and the entire model recalculates.

![Dashboard](docs/dashboard.png)
<!-- Open the xlsx in Excel, screenshot the Dashboard sheet, save as docs/dashboard.png -->

---

## 🔑 Key Features

| Area | What it does | Excel techniques |
|---|---|---|
| **Interactive Dashboard** | 15 KPIs (AUM, GNPA %, NNPA %, Provision Coverage, SMA stress, Wtd Avg Rate…) filtered by Region & Product dropdowns | Data validation, `SUMIFS` / `COUNTIFS` / `AVERAGEIFS`, dynamic filter-flag column, `SUMPRODUCT` weighted averages |
| **Asset Classification** | Buckets every loan into Current / SMA-0 / SMA-1 / SMA-2 / Sub-standard / Doubtful / Loss and applies provisioning | Approximate-match `INDEX`/`MATCH` against named ranges |
| **Loan Mechanics** | EMI, months elapsed, EMIs paid, outstanding principal, overdue amount per loan | `PMT`, `PV`, date arithmetic |
| **Top-10 NPA Leaderboard** | Largest NPA exposures for the current filter, updating live | `LARGE` + `ROW()` tie-break key + `INDEX`/`MATCH` |
| **Vintage (Cohort) Analysis** | GNPA % by disbursal quarter to detect underwriting drift | Heat-map colour scales, combo bar + line chart on dual axes |
| **EMI Calculator** | Full 360-month amortization schedule | `PMT`, `IPMT`, `PPMT`, `EDATE`, dynamic blank rows |
| **Controls** | Reconciliation check (bucket total = AUM), GNPA traffic light, input validation (CIBIL 300–900, DPD ≥ 0) | Conditional formatting (formula rules, data bars, colour scales), cell comments |

## 📁 Workbook Structure

```
Guide            → how to use the model, colour key
Dashboard        → KPIs, DPD / product / region breakdowns, Top-10 NPA, charts
Loans            → loan-level data (inputs in blue, formulas in black) as an Excel Table
Vintage          → cohort performance by disbursal quarter
EMI Calculator   → standalone EMI + amortization tool
Assumptions      → as-of date, provisioning table, CIBIL risk-grade bands (named ranges)
Branches         → branch → city → region master
```

## 📐 Model Logic

- **NPA:** loan overdue > 90 days (RBI definition). Sub-standard 91–365 DPD, Doubtful 366–730, Loss 731+.
- **Provisioning:** Standard 0.4% · Sub-standard 15% · Doubtful 40% (blended) · Loss 100%.
- **GNPA %** = NPA outstanding ÷ total outstanding. **NNPA** = GNPA − NPA provisions.
- **Provision Coverage Ratio** = NPA provisions ÷ GNPA.
- **Outstanding** = `PV(rate/12, tenure − EMIs paid, −EMI)`, where EMIs paid = months elapsed − ⌊DPD/30⌋.

All assumptions live on the **Assumptions** sheet — edit them and every number updates.

## 🚀 Run It

```bash
pip install -r requirements.txt
python src/build_workbook.py                 # 1,000 loans, seed 42
python src/build_workbook.py --rows 5000 --seed 7
```
Output: `output/Loan_Portfolio_Risk_Analytics.xlsx`. Open in Excel 2016+ (Microsoft 365 recommended).

## 💡 Sample Insights (default seed)

- Book of ~₹216 Cr AUM with **GNPA 2.7%** and **NNPA 1.9%**.
- Unsecured products (Personal, Business) run **10–14% GNPA** vs **<1%** for Home Loans — a classic secured/unsecured risk split.
- SMA buckets hold more exposure than current NPAs — an early-warning signal for future slippage.

## ⚠️ Disclaimer
All data is **synthetic** (seeded random). Provisioning rates are simplified from RBI IRACP norms for illustration and are not regulatory advice.

## 👤 Author

**Himanshu Kolhe**
- **MBA** — D.Y. Patil University, Navi Mumbai (2024–2026)
- **B.Tech** — JSPM's Rajarshi Shahu College of Engineering, Tathawade, Pune (2019–2023)

Aspiring **MIS Executive** combining an engineering background in logic and data with MBA training in business reporting.

**Skills shown in this project:**
- MS Excel: VLOOKUP, Pivot Tables, SUMIFS / COUNTIFS, IF / IFERROR, conditional formatting, charts
- Google Sheets: shared reports and dashboards
- Data verification: catching duplicates, missing codes and entry errors before reporting
- Consolidating data from Sales, Accounts and Stores into one MIS report

📧 himanshukolhe2002@gmail.com · 📍 Pune · 🔗 [LinkedIn](https://www.linkedin.com/in/himanshu-kolhe/)

*Open to MIS Executive / MIS Analyst / Reporting roles in Pune.*

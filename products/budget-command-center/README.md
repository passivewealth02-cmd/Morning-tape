# Budget & Money Command Center™ — The Complete Personal-Finance System

> Not a spreadsheet — a **complete give-every-dollar-a-job money system**. One
> premium **Google Sheets + printable PDF** command center for your whole
> financial life: a zero-based monthly budget (planned vs actual), income & bill
> trackers, a tagged expense log, savings goals, sinking funds, a debt snapshot,
> a net-worth tracker, a subscriptions audit, a 12-month year view and a no-spend
> challenge — all rolling up into one live Budget Health score.

| | |
| - | - |
| **Product** | Budget & Money Command Center™ |
| **Target** | Budget beginners & nerds alike · couples & families sharing money · debt-payoff & FIRE journeys · anyone living paycheck-to-paycheck · freelancers with variable income · new savers building a first fund |
| **Angle** | Give every dollar a job, crush debt & watch your net worth grow — your whole money life, organized. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $16 single · **$24 bundle** (Sheets + PDF) · $32 with the debt-payoff add-on · $79 coach / commercial license |

---

## Contents

```
products/budget-command-center/
├── README.md
├── Budget_Command_Center.xlsx    ← Google Sheets / Excel master (14 tabs)
├── Budget_Printables.pdf         ← 12-page print-ready pack (US Letter)
├── GOOGLE_SHEETS.md
├── BUILD_INSTRUCTIONS.md
├── ETSY_LISTING.md
└── build/
    ├── build_xlsx.py
    ├── build_pdf.py
    ├── build_marketing.py
    └── build_marketing_detail.py
```

---

## The 14-tab system

| # | Tab | # | Tab |
| - | --- | - | --- |
| 1 | Start Here | 8 | Sinking Funds |
| 2 | Dashboard | 9 | Debt Snapshot |
| 3 | Income | 10 | Net Worth |
| 4 | Monthly Budget | 11 | Subscriptions |
| 5 | Bills | 12 | Year View |
| 6 | Expense Log | 13 | No-Spend |
| 7 | Savings Goals | 14 | Settings |

## The 12 printable PDF pages

Monthly Budget (Zero-Based) · Bill Tracker · Expense Log · Savings Goals ·
Sinking Funds · Debt Snowball / Payoff · Net-Worth Worksheet · Subscriptions
Audit · Income Tracker · No-Spend Challenge · Year-at-a-Glance · Money Goals.

---

## Signature automation — a zero-based budget that balances itself

Give every dollar a job until **Left to Budget** hits `$0`:
`=Income-BudgetPlanTotal`. Each category shows **planned vs actual** with a
remaining column that flags red the moment it goes negative, and **Saved** is
pulled straight from the budget by group: `=SUMIF(BudgetGroup,"Savings",BudgetActual)`.

### The 12 dashboard KPIs
Income · Spent · Left to Budget · Saved · Savings Rate · Bills Paid · Net Worth ·
Total Debt · Savings Goals · Sinking Funds · Subscriptions · Health Score. The
**Budget Health Score** blends on-budget, savings rate, bills-on-time, emergency
fund, savings-goals progress and sinking-funds readiness into one 0–100% number,
next to a **Spending by Group** donut.

**Verified sample household** (The Bennett Household): Income **$5,200** · Spent
**$5,200** · Left to budget **$0** · Saved **$850** · Savings rate **16%** ·
Bills paid **80%** (8 of 10) · Net worth **$111,300** · Total debt **$30,200** ·
Savings goals **53%** · Sinking funds **$1,650** · Subscriptions **$80**/mo ·
**Health Score 80%**.

---

## Premium personal-finance-software design

- A **zero-based monthly budget** with planned-vs-actual and over-budget flags
- **Bill** & tagged **expense** trackers, **savings goals** with progress bars
- **Sinking funds** for big irregular costs; a **debt snapshot** sorted by rate
- A **net-worth** tracker (assets − liabilities) and a **subscriptions** audit
- A 12-month **year view** and a **no-spend** streak challenge
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A personal budgeting tool, not financial, tax or investment advice.** For big
> decisions, talk to a qualified professional.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Budget_Command_Center.xlsx
python3 build_pdf.py                        # -> ../Budget_Printables.pdf (+ page PNGs)
python3 build_marketing.py                 # -> ../marketing/01..06.png
python3 build_marketing_detail.py          # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

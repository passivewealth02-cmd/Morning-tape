# Budget & Money Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/budget-command-center/build
python3 build_xlsx.py      # -> ../Budget_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not financial advice" note.
2. **Income** sums every source (**$5,200**). **Monthly Budget** is zero-based —
   the header **Left to Budget** reads `=Income-BudgetPlanTotal` (**$0**), each
   category shows planned vs actual, and Remaining flags red when negative.
3. **Dashboard** fills 12 KPI cards + a Budget Health table + a spending donut.
   **Health Score 80%**. Income **$5,200** · Spent **$5,200** · Left **$0** ·
   Saved **$850** · Savings rate **16%** · Bills paid **80%**.
4. **Savings Goals** progress bars (`Saved/Target`) average **53%**; **Sinking
   Funds** total **$1,650**; **Debt Snapshot** totals **$30,200** with a
   rate-attack color scale.
5. **Net Worth** = assets − liabilities = **$111,300**. **Subscriptions** total
   **$80**/mo. Editing any input updates the dashboard live.
6. No broken cells; custom tables (Income, Monthly Budget, Savings Goals, Debt,
   Net Worth, Settings) start in column B.

> Note: uses `SUMIF`, `COUNTIF`, `COUNTA`, `AVERAGE`, `MIN`, `IFERROR` — opens in
> Google Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Budget_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter.
Twelve pages: monthly budget (zero-based), bill tracker, expense log, savings
goals, sinking funds, debt snowball/payoff, net-worth worksheet, subscriptions
audit, income tracker, no-spend challenge, year-at-a-glance and money goals.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard), everything-inside (14 tabs), the
zero-based monthly budget, the net-worth tracker, savings goals + bill tracker,
and the **12-page printables showcase**. Images 3–5 each show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "basic budget vs Command
Center", 09 run-your-money in 4 steps, 10 what's-included / who-it's-for /
guarantee. Ten images — fills all 10 Etsy slots. All headline numbers ($5,200
income · $0 left · $850 saved · 16% rate · 80% bills · $111,300 net worth ·
$30,200 debt · 53% goals · $1,650 sinking · $80 subs · 80% score) are verified
against the workbook.

---

## D. Etsy delivery package

```
Budget_Command_Center.xlsx         ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
Budget_Printables.pdf               ← 12-page print-ready pack
START_HERE.pdf                      ← onboarding quick-start
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| BMC-GS   | Google Sheets only | $16 |
| BMC-PDF  | Printable PDF only | $16 |
| BMC-BUNDLE | Sheets + PDF + Quick-Start | **$24** |
| BMC-PLUS | Bundle + debt-payoff add-on | $32 |
| BMC-PRO  | Coach / commercial-use license | $79 |

- **Sharp December–February peak** (New-Year budgeting resolutions) with a
  steady all-year tail. List by November and refresh tags in late December.
- Use all 10 photos + a walkthrough video. Lead photo = the feature-forward
  hero; the **zero-based budget** and the **net-worth tracker** are your
  strongest differentiators — most listings are a plain expense sheet.
- Cross-sell the **Debt Payoff Command Center** (snowball/avalanche) — budget →
  debt payoff is a natural bundle.

---

## F. Maintenance

- Edit the `INCOME_SOURCES`, `BUDGET`, `BILLS`, `SAVEGOALS`, `SINKING`, `DEBTS`,
  `ASSETS`, `SUBS` constants in `build_xlsx.py`; every KPI + the Health Score
  recompute. Add a category → add a `BUDGET` row (totals & donut follow).
- Printable pages live in `build_pdf.py` (one function per page). Keep
  `build_marketing.py` numbers in sync with the workbook.

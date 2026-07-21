# Debt Payoff Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/debt-command-center/build
python3 build_xlsx.py      # -> ../Debt_Command_Center.xlsx  (12 tabs)
```

The build **simulates** the snowball & avalanche payoff schedules month by month
and prints a summary. Expected (sample data):

```
Snowball: 42 mo, $9,908 interest, debt-free Jan 2030
Avalanche: 41 mo, $9,035 interest
Interest saved (avalanche): $873
Min-only interest: $24,010
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not financial advice" note.
2. **Debts** lists 6 debts; each shows **% Paid** (`(orig−balance)/orig`) with a
   data bar and an APR color scale. Totals: balance **$47,800**, min **$1,105**,
   original **$66,000**, **28%** paid.
3. **Payoff Plan** orders debts smallest-balance-first with projected payoff dates
   (Medical Bill → Dec 2026 … Student Loan → Jan 2030); focus debt = **Medical Bill**.
4. **Snowball vs Avalanche** shows months (**42** / **41**), total interest
   (**$9,908** / **$9,035**), debt-free dates and **$873** saved by avalanche.
5. **Dashboard** fills 12 KPI cards + a Payoff Momentum table + a balance bar
   chart. **Momentum 62%**. Monthly payment **$1,405** · debt-free **Jan 2030**.
6. No broken cells; custom tables (Debts, Payoff Plan, Interest Tracker, Net-Worth
   style panels, Settings) start in column B.

> Note: uses `SUMIF`, `COUNTIF`, `COUNTA`, `AVERAGE`, `MAX`, `MIN`, `INDEX`,
> `IFERROR` — opens in Google Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Debt_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter.
Twelve pages: debt list & snapshot, payoff plan, snowball vs avalanche, debt
payoff tracker, debt-free date worksheet, a **debt thermometer** to color,
interest-saved worksheet, extra-payment finder, payment log, balance-history
chart, milestones & wins, and debt-freedom goals.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard), everything-inside (12 tabs), the
debt-list engine, the snowball-vs-avalanche comparison, the payoff plan + payment
log, and the **12-page printables showcase**. Images 3–5 each show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "basic tracker vs Command
Center", 09 get-to-debt-free in 4 steps, 10 what's-included / who-it's-for /
guarantee. Ten images — fills all 10 Etsy slots. All headline numbers ($47,800
debt · 28% paid · $1,405/mo · Jan 2030 · 42 months · $9,908 interest · $873 saved
· $14,102 vs minimums · 62% momentum) are verified against the simulation.

---

## D. Etsy delivery package

```
Debt_Command_Center.xlsx           ← Google Sheets / Excel master (12 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
Debt_Printables.pdf                 ← 12-page print-ready pack
START_HERE.pdf                      ← onboarding quick-start
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| DPC-GS   | Google Sheets only | $16 |
| DPC-PDF  | Printable PDF only | $16 |
| DPC-BUNDLE | Sheets + PDF + Quick-Start | **$24** |
| DPC-PLUS | Bundle + budget add-on | $32 |
| DPC-PRO  | Coach / commercial-use license | $79 |

- **Sharp December–February peak** (New-Year debt-free resolutions) with a steady
  all-year tail. List by November and refresh tags in late December.
- Use all 10 photos + a walkthrough video. Lead photo = the feature-forward
  hero; the **snowball-vs-avalanche comparison** and the **projected debt-free
  date** are your strongest differentiators — most listings are a static tracker.
- Cross-sell the **Budget & Money Command Center** — budget → debt payoff is a
  natural bundle (the budget frees up the extra payment).

---

## F. Maintenance

- Edit the `DEBTS`, `EXTRA_PAYMENT`, `PAYMENTS`, `FOUND`, `MILESTONES`, `ACCEL`
  constants in `build_xlsx.py`; the `simulate()` engine + every KPI recompute.
  Add a debt → add a `DEBTS` row (the plan, dates & score follow) and re-run.
- Keep `build_marketing.py`'s KPIs and the payoff dates in sync with the printed
  simulation summary.

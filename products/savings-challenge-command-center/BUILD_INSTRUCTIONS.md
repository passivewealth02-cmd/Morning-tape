# Savings Challenge & Sinking Funds Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/savings-challenge-command-center/build
python3 build_xlsx.py      # -> ../Savings_Challenge_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not financial advice" note.
2. **Sinking Funds** sums 8 funds to **$9,000** target (`FundsTarget`) and **$6,000**
   saved (`FundsSaved`), with a **$750** monthly set-aside (`MonthlyNeed` = target ÷ 12)
   and each fund showing its own monthly number.
3. On-pace check: expected-to-date = $9,000 × 8 ÷ 12 = **$6,000** (`ExpectedToDate`), so
   **ON PACE = 100%** (`OnPace`). `FundCount` = 8.
4. **100 Envelope** shows the **$5,050** total (`EnvelopeTotal`); **52-Week** shows
   **$1,378** (`Week52Total`). Both track which are done via `SUMIF`.
5. **Emergency Fund** shows **$6,000 / $6,000** = 100% funded; **Savings Accounts**
   totals **$12,274** (`TotalSaved`); **Cash Envelopes** tracks loaded vs spent.
6. **Dashboard** fills 12 KPI cards + a Savings Health table + funds-filled &
   saved-by-month charts. **Savings Score 90%** (the 24-day streak against a 60-day goal
   is the honest weak dimension). No broken cells; custom tables start in column B.

> Note: uses `SUM`, `SUMIF`, `COUNTA`, `COUNTIF`, `AVERAGE`, `MIN`, `MAX`, `IF`,
> `IFERROR` — opens in Google Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Savings_Challenge_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter.
Twelve pages: sinking-funds worksheet, **the 100-envelope grid (10×10, numbered 1–100
with dollar values)**, **the 52-week tracker (8×7 grid)**, a savings thermometer, cash
envelope tracker, deposit log, goal countdown, no-spend tracker, savings streak,
emergency fund, monthly summary and a savings checklist.

> The 100-envelope grid and the thermometer are drawn programmatically by
> `grid_numbers()` — change the `cols`/`rows` arguments to reflow them.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with the funds-filled & saved charts),
everything-inside (14 tabs), the sinking-funds engine, all-funds table, the savings
engine (both), and the **12-page printables showcase**. Images 3–5 each show a different
tab.

**Four detailed images**: 07 feature spotlights, 08 "basic tracker vs Command Center",
09 start-saving in 4 steps, 10 what's-included / who-it's-for / works-with. Ten
images — fills all 10 Etsy slots. All headline numbers ($9,000 target · $6,000 saved ·
$750 monthly · 100% on pace · $6,000 EF · $12,274 total · $5,050 envelope · $1,378
52-week · 3 challenges · $520 this month · 24-day streak · 90% score) are verified
against the workbook.

---

## D. Etsy delivery package

```
Savings_Challenge_Command_Center.xlsx   ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt          ← "Make a Copy" link
Savings_Challenge_Printables.pdf         ← 12-page print-ready pack
START_HERE.pdf                           ← onboarding quick-start
```

---

## E. Pricing  (all SKUs are the same digital files — a product, never a service)

| SKU | What the buyer downloads | Price |
| --- | ------ | ----- |
| SCC-GS   | The Google Sheets / Excel file only | $15 |
| SCC-PDF  | The printable PDF only | $15 |
| SCC-BUNDLE | The spreadsheet + the printable PDF | **$22** |
| SCC-COMM | The same files + a commercial-use file license | $39 |

> Etsy Services-policy note: keep every listing a plain digital download. Do not
> advertise setup help, custom work, consultations, "done-for-you" builds, or "free
> updates / lifetime access" — offering a service (rather than a finished file) is
> what gets a listing removed and earns a strike.

- **The most shareable product in the catalogue.** Cash stuffing and the 100-envelope
  challenge are an active social-media movement — the printable grid gets screenshotted
  and reshared, which drives free traffic back to the listing.
- **Huge January peak**, plus a second spike every autumn as people scramble to fund
  Christmas. Price it lower than the business tools: this is a volume seller.
- Use all 10 photos + a walkthrough video. Lead photo = the hero; make **image 06 (the
  printables showcase)** prominent because the 100-envelope grid is the hook.
- Cross-sell the **Budget**, **Debt** and **Net Worth & FIRE** templates — same buyer,
  and a natural "Money Master Bundle" at $49–79.

---

## F. Maintenance

- Edit the `FUNDS`, `ENVELOPES`, `WEEKS_52`, `CHALLENGES`, `ACCOUNTS`, `DEPOSITS`,
  `GOALS`, `NOSPEND`, `CASH_ENVELOPES`, `MONTHS` tables and the `MONTHS_IN`,
  `MONTHLY_SAVE_GOAL`, `STREAK_GOAL`, `CHALLENGE_GOAL`, `FUNDS_GOAL_COUNT`, `EF_GOAL`,
  `EF_CURRENT`, `THIS_MONTH_SAVED`, `STREAK_DAYS` settings in `build_xlsx.py`; every
  KPI + the Savings Score recompute.
- **`MONTHS_IN` drives the on-pace check** — it's how far through the year the sample
  saver is. Keep it consistent with the `MONTHS` trend table.
- Keep `build_marketing.py`'s KPIs in sync with the workbook.

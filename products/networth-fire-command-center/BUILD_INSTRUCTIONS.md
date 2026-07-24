# Net Worth & FIRE Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/networth-fire-command-center/build
python3 build_xlsx.py      # -> ../Networth_FIRE_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not financial/tax advice"
   note.
2. **Assets** sums to **TOTAL ASSETS $530,000**; invested assets **$200,000**
   (`SUMIF` on "Investment"), cash **$20,000**. **Liabilities** sums to **$280,000**.
3. **Net Worth** = $530,000 − $280,000 = **$250,000** (named `NetWorth`); liquid net
   worth **$190,000**.
4. **FIRE Number** = $40,000 ÷ 4% = **$1,000,000** (`FIRENumber`); progress
   **25%**; coast number **$131,367** (`CoastNumber`); years to FI **10.5**
   (`YearsToFI`, via `NPER`).
5. **Income & Expenses** nets $80,000 − $40,000 = **$40,000 saved**, a **50%** savings
   rate (`SavingsRate`).
6. **Dashboard** fills 12 KPI cards + a FIRE Health table + progress-to-FIRE & net-
   worth-by-month charts. **FIRE Score 90%** (halfway-to-FI is the honest weak
   dimension). No broken cells; custom tables start in column B.

> Note: uses `SUM`, `SUMIF`, `COUNTIF`, `COUNTA`, `AVERAGE`, `MIN`, `NPER`, `IFERROR`
> — opens in Google Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Networth_FIRE_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter.
Twelve pages: net-worth worksheet, FIRE-number worksheet, asset list, liability list,
account list, contributions log, income & expenses, savings-rate log, coast &
projection, milestones, net-worth trend and a money checklist.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with the progress-to-FIRE & net-worth
charts), everything-inside (14 tabs), the FIRE-number engine, the coast-FIRE engine,
the FIRE engine (number + coast), and the **12-page printables showcase**. Images 3–5
each show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "basic tracker vs Command Center",
09 find-your-number in 4 steps, 10 what's-included / who-it's-for / works-with. Ten
images — fills all 10 Etsy slots. All headline numbers ($250k net worth · $1M FIRE
number · 25% progress · 50% savings rate · $131,367 coast · 10.5 years · $200k
invested · 90% score) are verified against the workbook.

---

## D. Etsy delivery package

```
Networth_FIRE_Command_Center.xlsx  ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
Networth_FIRE_Printables.pdf        ← 12-page print-ready pack
START_HERE.pdf                      ← onboarding quick-start
```

---

## E. Pricing  (all SKUs are the same digital files — a product, never a service)

| SKU | What the buyer downloads | Price |
| --- | ------ | ----- |
| NWF-GS   | The Google Sheets / Excel file only | $19 |
| NWF-PDF  | The printable PDF only | $19 |
| NWF-BUNDLE | The spreadsheet + the printable PDF | **$29** |
| NWF-COMM | The same files + a commercial-use file license | $49 |

> Etsy Services-policy note: keep every listing a plain digital download. Do not
> advertise setup help, custom work, consultations, "done-for-you" builds, or "free
> updates / lifetime access" — offering a service (rather than a finished file) is
> what gets a listing removed and earns a strike. A commercial-use license is a
> permission term attached to the *file* and is allowed; doing the work *for* the
> buyer is not.

- **Strong year-round demand** with a January (new-year money-reset) peak. FIRE and
  net-worth tools are among the best-selling finance templates on Etsy.
- Use all 10 photos + a walkthrough video that shows the *file*. Lead photo = the
  feature-forward hero; the **FIRE-number engine** and the **coast projection** are
  your strongest differentiators — most listings are just a net-worth total.
- Cross-sell the **Subscription & Bills Audit** and the **Freelancer** templates —
  same money-minded buyer.

---

## F. Maintenance

- Edit the `ASSETS`, `LIABILITIES`, `ACCOUNTS`, `CONTRIB`, `SAVINGS_LOG`,
  `MILESTONES`, `MONTHS`, `PROJ_AGES` constants and the `WITHDRAWAL_RATE`,
  `RETURN_RATE`, `CURRENT_AGE`, `RETIRE_AGE`, `ANNUAL_INCOME`, `ANNUAL_EXPENSES`,
  `CONTRIB_ANNUAL`, `SAVINGS_GOAL`, `EMERGENCY_FUND_GOAL`, `CONTRIB_GOAL`,
  `HALFWAY_GOAL` targets in `build_xlsx.py`; every KPI + the FIRE Score recompute.
  Everything is cross-linked — change your spending or your assets and the FIRE
  number, progress, coast and score follow.
- Keep `build_marketing.py`'s KPIs in sync with the workbook.

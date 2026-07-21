# Flip Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/flip-command-center/build
python3 build_xlsx.py      # -> ../Flip_Command_Center.xlsx  (17 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step quick guide + a "not financial
   advice" note.
2. **Deal Analyzer** takes ARV/purchase/rehab inputs and computes all-in cost,
   profit, cash-on-cash ROI, the 70%-rule MAO and a BUY/PASS verdict banner.
   Sample: All-In **$266,802**, Profit **$73,198**, ROI **77%**, MAO **$193,000
   → BUY**.
3. **Dashboard** fills 12 KPI cards + a Deal & Project Health table + a Rehab
   Planned-vs-Actual clustered bar chart. Deal Score **83%**.
4. **Rehab Budget** shows planned vs actual, % used and a red flag on any
   over-budget line (Demo, Electrical & Permits are over in the sample). Budget
   used **76%** ($34K of $45K).
5. Changing ARV, purchase price or the rehab budget re-runs the whole deal.
6. No broken cells; custom tables (Deal Analyzer, Rehab, Holding, Financing,
   Selling) start in column B.

> Note: uses `IF`, `IFERROR`, `MIN/MAX`, `SUMIF`, `COUNTIF`, `AVERAGE`, `TEXT` —
> opens in Google Sheets or Excel 2019/365. All deal outputs are live formulas.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Flip_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light job-site forms on white with a forest-green header band. 300 DPI, US
Letter. Twelve pages: deal analyzer worksheet, rehab budget, scope of work,
contractor & bid sheet, draw schedule, materials list, project timeline,
holding-cost worksheet, comps & ARV, selling/exit net sheet, punch list, and a
before/after photo log.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard), everything-inside (17 tabs), the
Deal Analyzer money shot (70% rule + BUY verdict), the rehab budget (overruns
flagged red), scope + timeline, and the **12-page printables showcase**. Images
3–5 each show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "basic spreadsheet vs
Command Center", 09 run-a-flip-in-4-steps, 10 what's-included / who-it's-for /
guarantee. Ten images — fills all 10 Etsy slots. All headline numbers ($340K ARV
· $185K buy · $45K rehab · $266,802 all-in · $73,198 profit · 77% ROI · $193K MAO
· BUY · Deal Score 83%) are verified against the workbook.

---

## D. Etsy delivery package

```
Flip_Command_Center.xlsx           ← Google Sheets / Excel master (17 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
Flip_Printables.pdf                 ← 12-page job-site binder
START_HERE.pdf                      ← onboarding quick-start
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| FCC-GS   | Google Sheets only | $19 |
| FCC-PDF  | Printable PDF only | $19 |
| FCC-BUNDLE | Sheets + PDF + Quick-Start | **$29** |
| FCC-PRO  | Investor / multi-deal edition | $39 |
| FCC-TEAM | Team / commercial-use license | $99 |

- **Higher-value buyer than a planner** — real-estate investors spend to protect
  a five-figure profit, so this supports a premium price and strong average order
  value. Steady year-round demand with a late-winter / spring bump (flip season).
- Two angles: **"know your number before you buy"** and **"the ultimate flipping
  system, offer to sold."** Bundle is the hero SKU; the investor/multi-deal
  edition and team license lift AOV.

---

## F. Maintenance

- Edit the `DEAL` dict and the `REHAB`, `SCOPE`, `CONTRACTORS`, `PAYMENTS`,
  `MATERIALS`, `PHASES`, `HOLDING`, `COMPS` constants in `build_xlsx.py`; every
  KPI, the deal outputs and the Deal Score recompute. The Deal Analyzer inputs
  are named cells — change them and the whole model follows.
- Printable pages live in `build_pdf.py` (one function per page). Keep
  `build_marketing.py` numbers in sync with the workbook (they're a manual mirror).

# Dividend Wealth Builder — Build Instructions

Two reproduction paths: **(A)** rebuild the Excel `.xlsx` from source,
**(B)** assemble the Google Sheets version from scratch.

---

## A. Excel build (from `build/build_xlsx.py`)

### Requirements

- Python 3.10+
- `openpyxl >= 3.1.5`

```bash
pip install openpyxl
```

### Build

```bash
cd products/dividend-wealth-builder/build
python3 build_xlsx.py
# → ../Dividend_Wealth_Builder.xlsx
```

The script is fully deterministic — same code, same workbook every run.

### Verifying

Open `Dividend_Wealth_Builder.xlsx` in Excel 365 or 2021+. On open:

1. **Dashboard** tab populates KPI cards with live values from sample
   holdings.
2. **Holdings** rows 5–19 are pre-filled. Adding a row anywhere from row
   20 to row 64 auto-calculates Total Value / Annual Income / Yield via
   formulas already in place.
3. **Settings → C5–C11** edits propagate through every KPI, projection,
   and chart instantly.

---

## B. Google Sheets build

Follow `GOOGLE_SHEETS.md` end-to-end. Order matters:

1. Create blank sheet → rename to **Dividend Wealth Builder**.
2. Add tabs in this order: `Dashboard`, `Holdings`, `Calculations`,
   `Settings`, `Projections`.
3. Build **Settings** first (section 1) — defines every named range the
   other sheets rely on.
4. Build **Holdings** (section 2) — sample data, validation,
   conditional formatting.
5. Build **Calculations** (section 3) — aggregates and sector breakdown.
6. Build **Projections** (section 4) — FIRE metrics + 30-year table.
7. Build **Dashboard** (section 5) — KPI cards + 6 embedded charts.
8. Paste the Apps Script in section 6 → run `installTriggers`.
9. Apply brand palette (section 7).

---

## C. Etsy listing structure (delivery package)

Bundle ships as a single `.zip`:

```
Dividend_Wealth_Builder.xlsx                ← Excel master
GOOGLE_SHEETS_TEMPLATE_LINK.txt             ← "Make a copy" link to the
                                              Google Sheets edition
START_HERE.pdf                              ← One-page walkthrough
THANK_YOU.pdf                               ← Brand-aligned thank-you
                                              card with support email
```

Listing photos (6 required):

1. Dashboard hero shot — full-bleed KPI row + charts.
2. Holdings sheet — alternating tan rows, validation dropdowns visible.
3. Sector allocation pie + monthly bar side-by-side.
4. Projections wealth-curve chart with FIRE callouts.
5. Settings panel — branded inputs, dropdowns.
6. Mobile preview (Google Sheets app open).

---

## D. Pricing

| SKU | Format | Etsy price |
| --- | ------ | ---------- |
| DWBS-EX | Excel only | $19 |
| DWBS-GS | Google Sheets only | $19 |
| DWBS-BUNDLE | Excel + Google Sheets + Quick-Start PDF | **$29** |
| DWBS-PRO | Bundle + 30-min onboarding call | $79 |

---

## E. Update / maintenance protocol

- **Quarterly:** refresh sample tickers + prices in `SAMPLE_HOLDINGS`
  (build_xlsx.py:131) and rerun the build.
- **Annual:** bump FIRE multiplier if the 25× rule changes
  (`Projections!C5`).
- **Bug reports → patch:** edit `build_xlsx.py`, rerun, replace the
  shipped `.xlsx` and the Google Sheets master template. Bump the
  footer version on the Dashboard.

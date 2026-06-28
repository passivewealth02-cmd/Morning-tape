# Back-to-School Command Center — Build Instructions

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
cd products/back-to-school-command-center/build
python3 build_xlsx.py
# → ../Back_To_School_Command_Center.xlsx
```

Deterministic — same script, same workbook every run.

### Verifying

Open in Excel 365 / 2021+. On open:

1. **Dashboard** populates all 8 KPI cards from live data on
   Settings + Supplies + Budget + Assignments.
2. **Days Until School** counts down from `TODAY()` — change
   `Settings!C6` to test.
3. **Supplies** rows 5–26 are pre-filled; add a row anywhere up to row
   64 and Remaining / completion KPIs update automatically.
4. **Budget!C5:C13** auto-sums actual spend from Supplies per category;
   change `Supplies!H5` (any actual price) → Budget Spent and Dashboard
   KPIs both update.
5. **Assignments** highlights overdue rows in red, complete in mint,
   due-this-week in tan.

---

## B. Google Sheets build

Follow `GOOGLE_SHEETS.md` in this exact order — later tabs depend on
earlier named ranges:

1. Create the blank sheet; add all 12 tabs.
2. **Settings** (section 1) — every named range lives here.
3. **Supplies** (section 2) — `ARRAYFORMULA` Remaining column +
   validation + conditional formats.
4. **Budget** (section 3) — `SUMIFS` against `SupActual`, totals row,
   two charts.
5. **Assignments** (section 4) — Days-Until column + 3 conditional
   format rules.
6. **Calendar** (section 5) — past / upcoming conditional formats.
7. **Students, Clothing, Schedule, Extracurricular, Meals, Emergency**
   — pure data tables; copy from the workbook templates.
8. **Dashboard** (section 6) — both KPI rows + embedded charts +
   navigation chips.
9. Paste Apps Script (section 7) → run `installTriggers`.
10. Apply brand palette (section 8).

---

## C. Etsy listing structure (delivery package)

Bundle ships as a single `.zip`:

```
Back_To_School_Command_Center.xlsx        ← Excel master
GOOGLE_SHEETS_TEMPLATE_LINK.txt           ← "Make a copy" link to
                                            the Google Sheets edition
START_HERE.pdf                            ← 1-page walkthrough
THANK_YOU.pdf                             ← brand thank-you card
                                            + support email
```

Listing photos (6 required):

1. Hero thumbnail (Driver-Budget format).
2. Dashboard close-up.
3. Supplies tracker (validation visible).
4. Budget breakdown + charts.
5. Class schedule (color-coded by subject).
6. Mobile preview (Google Sheets app).

---

## D. Pricing

| SKU | Format | Etsy price |
| --- | ------ | ---------- |
| BTSCC-EX     | Excel only | $19 |
| BTSCC-GS     | Google Sheets only | $19 |
| BTSCC-BUNDLE | Excel + Google Sheets + Quick-Start PDF | **$29** |
| BTSCC-FAMILY | Bundle + 30-min onboarding call | $79 |

Per the PRD this product can hold a **$20 – $40** Etsy price point.
The bundle at $29 hits the upper-middle of that band.

---

## E. Update / maintenance protocol

- **Annually (July):** bump the default `FirstDayOfSchool` and
  `SchoolYear` in `build_xlsx.py` (constants near
  `build_settings()`), rerun the build, replace the shipped `.xlsx`
  + Google Sheets master template.
- **Bug reports → patch:** edit `build_xlsx.py`, rerun, swap the
  `.xlsx` in the Etsy listing, bump the Dashboard footer version.
- **New sheet requests:** add a new `build_<name>()` function and slot
  it into the `order` list in `main()`. Brand styles + dropdown
  validation are already centralized.

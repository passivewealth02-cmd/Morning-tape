# Soccer Mom Command Center™ — Build Instructions

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
cd products/soccer-mom-command-center/build
python3 build_xlsx.py
# → ../Soccer_Mom_Command_Center.xlsx
```

Deterministic — same script, same workbook every run.

### Verifying

Open in Excel 365 / 2021+. On open:

1. **Dashboard** populates 8 KPI cards from live data on Settings +
   Schedule + Practices + Budget + Equipment + Tournaments.
2. **Games This Month** uses `MONTH()` / `YEAR()` over `GameDates`
   named range — add or shift a row on Schedule and the count updates.
3. **Practices This Week** uses date-window `SUMPRODUCT` — edit
   `Settings!C13` is recomputed via `TODAY()` automatically.
4. **Equipment Needed** counts `EquipPurchased <> "Yes"`. Toggle a row
   from `No` → `Yes` on Equipment and the KPI ticks down.
5. **Schedule** rows in the next 7 days highlight mint; past games go
   muted gray; Win / Loss tint the Result column.
6. **Mileage** Fuel and Total auto-compute via `FuelPerMile` from
   Settings.

---

## B. Google Sheets build

Follow `GOOGLE_SHEETS.md` in this exact order — later tabs depend on
earlier named ranges:

1. Create the blank sheet; add all 14 tabs.
2. **Settings** (section 1) — every named range lives here.
3. **Schedule** (section 2) — Days-Out `ARRAYFORMULA`, validation,
   upcoming / past / win-loss conditional formats.
4. **Practices** (section 3) — attendance dropdown + 3 colored states.
5. **Budget** (section 4) — variance / % spent / totals + 2 charts.
6. **Equipment** (section 5) — Replace By window + Yes/No formats.
7. **Tournaments** (section 6) — upcoming 30-day highlight.
8. **Players, Carpool, Roster, Meals, Packing, Communication** —
   pure data tables; copy from the workbook templates.
9. **Mileage** (section 7) — `ARRAYFORMULA` Fuel + Total.
10. **Dashboard** (section 8) — both KPI rows + nav chips + 5 charts.
11. Paste Apps Script (section 9) → run `installTriggers`.
12. Apply brand palette (section 10).

---

## C. Etsy listing structure (delivery package)

Bundle ships as a single `.zip`:

```
Soccer_Mom_Command_Center.xlsx        ← Excel master
GOOGLE_SHEETS_TEMPLATE_LINK.txt       ← "Make a copy" link to the
                                        Google Sheets edition
START_HERE.pdf                        ← 1-page walkthrough
THANK_YOU.pdf                         ← brand thank-you card
                                        + support email
```

Listing photos (6 required):

1. Hero thumbnail (Driver-Budget format).
2. Dashboard close-up.
3. Season schedule with countdown column.
4. Budget breakdown + charts.
5. Tournament planner / packing checklist.
6. Mobile preview (Google Sheets app).

---

## D. Pricing

| SKU | Format | Etsy price |
| --- | ------ | ---------- |
| SMCC-EX     | Excel only | $24 |
| SMCC-GS     | Google Sheets only | $24 |
| SMCC-BUNDLE | Excel + Google Sheets + Quick-Start PDF | **$34** |
| SMCC-CLUB   | Bundle + 30-min coach onboarding call | $89 |
| SMCC-TEAM   | Team license (up to 15 families) | $149 |

The bundle at $34 hits the premium end of the family-planner band on
Etsy. The Team SKU is the new high-ticket variant — team managers
become repeat customers.

---

## E. Update / maintenance protocol

- **Pre-season (Jul / Jan):** bump `SeasonStart` / `SeasonEnd` /
  `Season` defaults in `build_xlsx.py` (constants near
  `build_settings()`), rerun the build, replace the shipped `.xlsx` +
  Google Sheets master.
- **Bug reports → patch:** edit `build_xlsx.py`, rerun, swap the
  `.xlsx` on the Etsy listing, bump the Dashboard footer version.
- **Adding sheets** (e.g., `Stats`, `Highlights`): add a new
  `build_<name>()` function and slot the sheet into the `order` list
  in `main()`. Brand styles + dropdown validation are already
  centralized.
- **Coach license variant:** add a `build_coach_extras(wb)` that
  inserts a `Lineups` sheet + tactical drawing board (image placeholder
  rows). Same generator, different SKU.

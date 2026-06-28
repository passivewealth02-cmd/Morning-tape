# Wedding Command Center™ — Build Instructions

The flagship of the Spreadsheet Empire. 32 sheets. Two reproduction
paths below.

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
cd products/wedding-command-center/build
python3 build_xlsx.py
# → ../Wedding_Command_Center.xlsx  (32 sheets)
```

Deterministic — same script, same workbook every run.

### Verifying

Open in Excel 365 / 2021+. On open:

1. **Dashboard** populates 10 KPI cards across two rows + 4 charts
   (Budget Breakdown, Budget vs Actual, RSVP Progress, Readiness by
   Dimension).
2. **Days Until Wedding** counts down from `Settings!C6`.
3. **Budget** flows into the dashboard; change a `Deposit` cell and
   Payments Outstanding + Budget Remaining update.
4. **Guests** RSVP column drives RSVP %, Guest Count, Tables Filled,
   meal counts (RSVP sheet), and Guest Completion (Analytics).
5. **Analytics** computes a weighted **Wedding Readiness Score** (big
   % panel) from six health dimensions.
6. **Seating** flags over-capacity tables red, full tables mint, VIP
   tables gold.
7. **VisionBoard** shows 16 blush image-placeholder tiles ready for
   Insert ▸ Pictures ▸ Place in Cell.

---

## B. Google Sheets build

Follow `GOOGLE_SHEETS.md`. Critical order:

1. Create all 32 tabs.
2. **Settings** first (every named range).
3. Define the cross-sheet named ranges (GOOGLE_SHEETS.md §2).
4. Build data sheets: Budget → Payments → Guests → Vendors → Timeline →
   Checklist (these feed the dashboard).
5. **RSVP**, **Seating**, **Analytics** (aggregation sheets).
6. **Dashboard** last — KPIs + charts + hyperlinked nav chips.
7. **VisionBoard** — image-in-cell tiles.
8. Apps Script triggers (§6) → run `installTriggers`.
9. Apply the luxury palette (§7).

---

## C. Sheet inventory (32)

| # | Sheet | # | Sheet |
| - | ----- | - | ----- |
| 1 | Executive Dashboard | 17 | Ceremony Planner |
| 2 | Couple Profile | 18 | Reception Planner |
| 3 | Master Timeline | 19 | Menu Planner |
| 4 | Master Checklist | 20 | Music Planner |
| 5 | Budget Command Center | 21 | Floral Planner |
| 6 | Payment Schedule | 22 | Decor Planner |
| 7 | Savings Planner | 23 | Photography Shot List |
| 8 | Vendor CRM | 24 | Honeymoon Planner |
| 9 | Vendor Comparison | 25 | Registry Tracker |
| 10 | Contract Tracker | 26 | Gift Tracker |
| 11 | Guest CRM | 27 | Thank-You Tracker |
| 12 | RSVP Dashboard | 28 | Vision Board |
| 13 | Seating Planner | 29 | Packing Checklist |
| 14 | Bridal Party | 30 | Wedding Day Command Center |
| 15 | Dress & Attire | 31 | Analytics Dashboard |
| 16 | Beauty Timeline | 32 | Settings |

---

## D. Etsy listing structure (delivery package)

```
Wedding_Command_Center.xlsx          ← Excel master (32 sheets)
GOOGLE_SHEETS_TEMPLATE_LINK.txt      ← "Make a copy" link
START_HERE.pdf                       ← onboarding walkthrough
HOW_TO_ADD_PHOTOS.pdf                ← vision-board image guide
THANK_YOU.pdf                        ← luxury thank-you card
```

Listing photos (10 — use Etsy's full allotment for the flagship):

1. Hero thumbnail (Driver-Budget format).
2. Executive dashboard close-up.
3. Budget command center + charts.
4. Guest CRM + RSVP dashboard.
5. Vendor CRM / comparison.
6. Seating planner.
7. Master timeline.
8. Vision board (image tiles).
9. Wedding-day run of show.
10. Mobile preview.

---

## E. Pricing — flagship tier

| SKU | Format | Etsy price |
| --- | ------ | ---------- |
| WCC-EX      | Excel only | $39 |
| WCC-GS      | Google Sheets only | $39 |
| WCC-BUNDLE  | Excel + Google Sheets + onboarding PDFs | **$49** |
| WCC-PLANNER | Bundle + planner/coordinator white-label license | $129 |
| WCC-PRO     | Bundle + 45-min setup call | $99 |

This is the highest-priced product in the collection — the 32-sheet
depth, readiness score, and vision board justify the premium.

---

## F. Update / maintenance protocol

- **Per couple:** the buyer just edits `Settings` — date, budget,
  guests, tables. No formula edits needed.
- **Seasonal refresh:** update sample vendor names / dates in
  `build_xlsx.py` if screenshots feel dated; rerun.
- **New sheet:** add a `build_<name>()` (or extend `build_remaining`
  via the `build_generic` helper) and append to the `order` list in
  `main()`. Brand styles, dropdowns, and conditional formats are
  centralized.
- **White-label (WCC-PLANNER):** swap the monogram + couple names on
  Couple/Dashboard with the planner's branding; everything else is
  reusable.

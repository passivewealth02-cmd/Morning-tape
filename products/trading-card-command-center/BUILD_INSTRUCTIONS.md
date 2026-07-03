# Trading Card Collection Command Center™ — Build Instructions

---

## A. Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/trading-card-command-center/build
python3 build_xlsx.py      # -> ../Trading_Card_Command_Center.xlsx  (13-tab system + Welcome)
```

### Verifying
1. **Welcome** shows the intro + a Start-Here quick guide.
2. **Collection Dashboard** fills 12 KPI cards + 4 charts (value by rarity,
   value over time, value by set, grading before-vs-after).
3. **Master Collection** — Total 146 cards, $2,840 value, $1,124 cost,
   +$1,716 P/L (+153%); every row computes qty × value with a data-bar.
4. **Purchases** total reconciles to the collection's cost ($1,124).
   **Sales / Trades** net out fees & swap values (+$18 trade net).
5. **Grading Center** — 4 returned (PSA 10 / BGS 9.5 / PSA 9 / PSA 8, +$997
   value added), 2 pending. **Card Vault** shows the photo-upload guide and
   8 slots; the **Photos on File** KPI reads 67% (12 of 18).
6. **Analytics** blends a **Collection Health Score** (77%). No broken cells.

> Note: uses `SUMIF`, `SUMPRODUCT`, `COUNTIF`, `MAX` — opens in Excel 2019/365
> or Google Sheets.

---

## B. Google Sheets build

Follow `GOOGLE_SHEETS.md`: build **Settings** first, define the cross-sheet
named ranges, then the Master Collection and trackers, then the **Dashboard**
+ **Analytics**. The Card Vault works with Insert ▸ Image or `=IMAGE("url")`.

---

## C. Marketing images

```bash
cd build && python3 build_marketing.py     # -> ../marketing/01..06.png
```

Six 2000×2000 PNGs, rendered as dense app screenshots (sidebar of all 13 tabs
+ the real computed KPI numbers + fully populated tables/charts): hero,
everything-inside (13-tab showcase), collection dashboard, master collection,
grading + trades, and mobile.

---

## D. Etsy delivery package

```
Trading_Card_Command_Center.xlsx   ← Excel master (13-tab system + Welcome)
GOOGLE_SHEETS_TEMPLATE_LINK.txt    ← "Make a Copy" link
START_HERE.pdf                     ← onboarding quick-start
THANK_YOU.pdf                      ← brand thank-you card
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| TCC-EX     | Excel only | $19 |
| TCC-GS     | Google Sheets only | $19 |
| TCC-BUNDLE | Excel + Google Sheets + Quick-Start PDF | **$29** |
| TCC-PLUS   | Bundle + collector's grading & selling playbook | $39 |
| TCC-PRO    | Card shop / content-creator license | $79 |

- **Evergreen hobby with investment energy** — bumps around **new set
  releases**, **Q4 (gifts for collectors)** and grading-price news cycles.
- Lean into two angles: **"know what your collection is worth"** (investor
  intent) and **"finally organized"** (parent/hobbyist intent).
- Game-agnostic positioning widens the market: Pokémon, MTG, Yu-Gi-Oh!,
  One Piece, Lorcana, sports cards.

---

## F. Maintenance

- Edit sample cards in `build_xlsx.py` and rerun; values, P/L, health and all
  dashboard KPIs recompute automatically.
- Brand styles, dropdowns, KPI cards & summary tables are centralized.
- New sheet: add a builder and slot it into the `order` list in `main()`.
- Marketing numbers in `build_marketing.py` (`KPIS` list + content functions)
  should be kept in sync with the workbook's sample data.

# Bar & Pub Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/bar-command-center/build
python3 build_xlsx.py      # -> ../Bar_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not accounting advice" note.
2. **Pour Cost** costs an Old Fashioned to **$2.20** (spirit + mixer + garnish),
   linked into the Drink Menu via `=OldFashionedCost`.
3. **Drink Menu** shows pour cost, price, pour-cost % & margin. Avg pour cost
   **$1.61**, avg price **$9.62**, avg margin **$8.02**, overall pour cost **16%**;
   top seller (by revenue) = Draft Beer.
4. **Keg & Draft** rolls up to avg profit/keg **$571**; **Inventory Variance**
   totals **6.1%**; **Happy Hour** margin **79%**. Weekly sales **$16,940**,
   **1,870** units.
5. **Dashboard** fills 12 KPI cards + a Bar Health table + a sales-by-day chart.
   **Bar Score 90%**.
6. No broken cells; custom tables (Pour Cost, Drink Menu, etc.) start in column B.

> Note: uses `SUMPRODUCT`, `COUNTIF`, `COUNTA`, `AVERAGE`, `MAX`, `MATCH`,
> `INDEX`, `MIN`, `IFERROR` — opens in Google Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Bar_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter.
Twelve pages: pour cost card, drink menu, keg & draft log, inventory variance,
liquor count sheet, happy hour planner, weekly sales log, waste & spill log,
ordering sheet, cash & tips, events & big tabs, and an open/close checklist.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with variance & sales charts),
everything-inside (14 tabs), the drink menu (pour-cost % & margin), the inventory
variance sheet, the pour-cost engine, and the **12-page printables showcase**.
Images 3–5 each show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "basic sales log vs Command
Center", 09 price-your-bar in 4 steps, 10 what's-included / who-it's-for /
guarantee. Ten images — fills all 10 Etsy slots. All headline numbers (8 drinks ·
$1.61 pour cost · 16% pour cost % · $16,940 weekly sales · 1,870 units · $571
profit/keg · 6.1% variance · 90% score) are verified against the workbook.

---

## D. Etsy delivery package

```
Bar_Command_Center.xlsx            ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
Bar_Printables.pdf                  ← 12-page print-ready pack
START_HERE.pdf                      ← onboarding quick-start
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| BPC-GS   | Google Sheets only | $19 |
| BPC-PDF  | Printable PDF only | $19 |
| BPC-BUNDLE | Sheets + PDF + Quick-Start | **$29** |
| BPC-PLUS | Bundle + pour-cost add-on | $39 |
| BPC-PRO  | Commercial / multi-location license | $99 |

- **Steady all-year demand** with peaks in Q4 (holiday events season) and January
  (new-year business planning). Priced above consumer planners — this saves an
  over-pouring bar real money.
- Use all 10 photos + a walkthrough video. Lead photo = the feature-forward hero;
  the **pour-cost engine** and **inventory variance** are your strongest
  differentiators — most listings are a plain sales log.
- Cross-sell the **Restaurant**, **Café** and **Food Truck** products — same
  buyer, natural "food & beverage business" bundle.

---

## F. Maintenance

- Edit the `POURCOST`, `DRINKS`, `KEGS`, `VARIANCE`, `LIQUOR`, `HAPPYHOUR`,
  `WEEKLY`, `WASTE`, `ORDERING`, `CASH`, `EVENTS` constants in `build_xlsx.py`;
  every KPI + the Bar Score recompute. Add a drink → add a `DRINKS` row (pour-cost
  %, margin & totals follow).
- Keep `build_marketing.py`'s KPIs in sync with the workbook.

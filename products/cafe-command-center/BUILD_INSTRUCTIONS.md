# Cafe & Coffee Shop Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/cafe-command-center/build
python3 build_xlsx.py      # -> ../Cafe_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not accounting advice" note.
2. **Cup Cost** costs a Latte to the cup (beans, milk, cup, lid, sleeve) = **$1.18**,
   which links into the Menu Board Latte row via `=LatteCost`.
3. **Menu Board** shows beverage-cost % & margin per item; overall beverage cost
   **22%**. Avg cup cost **$1.13**, avg price **$5.05**, avg margin **$3.92**.
4. **Daypart Sales** totals **$2,769** across **352** transactions (avg ticket
   **$7.87**); top daypart = Morning rush.
5. **Labor & Prime** shows labor **28%** and **prime cost 50%** (bev + labor).
   **Dashboard** fills 12 KPI cards + a Café Health table + a daypart doughnut.
   **Café Score 91%**.
6. No broken cells; custom tables (Cup Cost, Menu Board, Daypart, etc.) start in column B.

> Note: uses `SUMPRODUCT`, `COUNTIF`, `COUNTA`, `AVERAGE`, `MAX`, `MATCH`,
> `INDEX`, `MIN`, `IFERROR` — opens in Google Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Cafe_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter.
Twelve pages: cup-cost card, menu board, daypart & weekly sales logs, labor &
prime cost, bean & milk usage, inventory & par, waste log, ordering sheet, cash &
tips, an open/close checklist, and a regulars & loyalty page.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with daypart & weekly charts),
everything-inside (14 tabs), the menu board with beverage-cost %, the labor &
prime-cost view, the cup-cost engine, and the **12-page printables showcase**.
Images 3–5 each show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "basic tracker vs Command
Center", 09 dial-it-in in 4 steps, 10 what's-included / who-it's-for / guarantee.
Ten images — fills all 10 Etsy slots. All headline numbers (10 items · $1.13 cup
cost · 22% bev · 28% labor · $2,769 daily · $7.87 ticket · 50% prime · 91% score)
are verified against the workbook.

---

## D. Etsy delivery package

```
Cafe_Command_Center.xlsx           ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
Cafe_Printables.pdf                 ← 12-page print-ready pack
START_HERE.pdf                      ← onboarding quick-start
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| CFC-GS   | Google Sheets only | $19 |
| CFC-PDF  | Printable PDF only | $19 |
| CFC-BUNDLE | Sheets + PDF + Quick-Start | **$29** |
| CFC-PLUS | Bundle + recipe-costing add-on | $39 |
| CFC-PRO  | Multi-location / commercial license | $99 |

- **Steady all-year B2B demand** with a January (new-year planning) bump. Priced
  above consumer planners — this saves an owner real money.
- Use all 10 photos + a walkthrough video. Lead photo = the feature-forward
  hero; the **cup-cost engine** and **prime cost** view are your strongest
  differentiators — most listings are a plain sales log.
- Cross-sell the **Recipe Costing** and **Bakery / Bar** products — same buyer,
  natural "food & beverage" bundle.

---

## F. Maintenance

- Edit the `LATTE`, `MENU`, `DAYPART`, `WEEKLY`, `LABOR`, `USAGE`, `INVENTORY`,
  `WASTE`, `ORDERING`, `CASHTIPS`, `REGULARS` constants in `build_xlsx.py`; every
  KPI + the Café Score recompute. Add a drink → add a `MENU` row.
- Keep `build_marketing.py`'s KPIs in sync with the workbook.

# Food Truck Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/foodtruck-command-center/build
python3 build_xlsx.py      # -> ../Food_Truck_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not legal advice" note.
2. **Events** nets out each gig (**sales − food − fuel − fee − staff**). Totals:
   sales **$15,230**, food **$4,429**, profit **$7,566**; top event by net =
   Corporate Catering ($1,695).
3. **Commissary & Overhead** totals **$2,140**/mo; **Break-Even** = overhead ÷
   avg net = **2.3 events**.
4. **Dashboard** fills 12 KPI cards + a Truck Health table + a net-profit-by-event
   bar chart. **Truck Score 93%**. Avg profit **$946**/event, food cost **29%**.
5. **Permits** shows **4 of 5** current; **Bookings** counts confirmed gigs;
   **Inventory & Par** flags low stock red.
6. No broken cells; custom tables (Menu, Events, Break-Even, etc.) start in column B.

> Note: uses `COUNTIF`, `COUNTA`, `AVERAGE`, `MAX`, `MATCH`, `INDEX`, `MIN`,
> `IFERROR` — opens in Google Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Food_Truck_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter.
Twelve pages: event P&L sheet, break-even worksheet, daily sales log, menu & cost
card, inventory & par, prep list, fuel & mileage log, permit tracker, shopping
list, bookings calendar, cash & tips reconciliation, and a monthly P&L summary.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with net-profit chart),
everything-inside (14 tabs), the event P&L, the break-even calculator, the menu &
bookings, and the **12-page printables showcase**. Images 3–5 each show a
different tab.

**Four detailed images**: 07 feature spotlights, 08 "basic tracker vs Command
Center", 09 run-the-numbers in 4 steps, 10 what's-included / who-it's-for /
guarantee. Ten images — fills all 10 Etsy slots. All headline numbers (8 events ·
$15,230 sales · $7,566 profit · $946/event · 29% food cost · 2.3 break-even ·
$2,140 overhead · 93% score) are verified against the workbook.

---

## D. Etsy delivery package

```
Food_Truck_Command_Center.xlsx     ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
Food_Truck_Printables.pdf           ← 12-page print-ready pack
START_HERE.pdf                      ← onboarding quick-start
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| FTC-GS   | Google Sheets only | $19 |
| FTC-PDF  | Printable PDF only | $19 |
| FTC-BUNDLE | Sheets + PDF + Quick-Start | **$29** |
| FTC-PLUS | Bundle + recipe-costing add-on | $39 |
| FTC-PRO  | Multi-truck / commercial license | $99 |

- **Spring–summer peak** (festival & event season) with steady year-round demand.
  List by March; refresh tags before festival season.
- Use all 10 photos + a walkthrough video. Lead photo = the feature-forward
  hero; the **event P&L** and **break-even** are your strongest differentiators —
  most listings are a plain sales log, not a per-gig profit + break-even system.
- Cross-sell the **Recipe Costing** and other food-service products — same buyer,
  natural "food business" bundle.

---

## F. Maintenance

- Edit the `MENU`, `EVENTS`, `OVERHEAD`, `DAILY`, `INVENTORY`, `FUEL`, `PERMITS`,
  `SUPPLIES`, `BOOKINGS`, `CASHTIPS` constants in `build_xlsx.py`; every KPI + the
  Truck Score recompute. Add a gig → add an `EVENTS` row (net & totals follow).
- Keep `build_marketing.py`'s KPIs and net-profit figures in sync with the workbook.

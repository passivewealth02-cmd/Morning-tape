# Catering Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/catering-command-center/build
python3 build_xlsx.py      # -> ../Catering_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not accounting advice" note.
2. **Plate Costing** costs a Plated Dinner to **$14.00** per head (protein + sides
   + dessert + overhead), linked into Menu Packages via `=PlatedCostHead`.
3. **Menu Packages** shows cost/head, price/head, margin & food-cost % on 8
   packages. The Plated Dinner is 29% food cost at $48/head.
4. **Event Quotes** turns each quote into a full event P&L — guests × package price
   + service, minus food, staff & rentals. Total revenue **$21,560**, food cost
   **26%**, avg event margin **32%**, labor **25%**; top package = Wedding Premium
   ($7,550/event).
5. **Dashboard** fills 12 KPI cards + a Catering Health table + a revenue-by-event
   chart. **Catering Score 90%**. Waste **2.0%**.
6. No broken cells; custom tables (Plate Costing, Menu Packages, Event Quotes,
   etc.) start in column B.

> Note: uses `INDEX`, `MATCH`, `SUMPRODUCT`, `COUNTIF`, `COUNTA`, `AVERAGE`, `MAX`,
> `MIN`, `IFERROR` — opens in Google Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Catering_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter.
Twelve pages: plate cost card, menu package price list, event quote sheet, event
run sheet, staffing sheet, rentals & equipment, bookings calendar, inventory & par,
waste log, ordering sheet, cash & deposits, and a client contact sheet.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with food-cost & revenue charts),
everything-inside (14 tabs), the menu packages (margin per head), the event quote
P&L, the plate-costing engine, and the **12-page printables showcase**. Images 3–5
each show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "basic price list vs Command
Center", 09 price-your-catering in 4 steps, 10 what's-included / who-it's-for /
guarantee. Ten images — fills all 10 Etsy slots. All headline numbers (6 events ·
$14.00 cost/head · 26% food cost · $21,560 revenue · $3,593 avg event · 32% margin
· 90% score) are verified against the workbook.

---

## D. Etsy delivery package

```
Catering_Command_Center.xlsx       ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
Catering_Printables.pdf             ← 12-page print-ready pack
START_HERE.pdf                      ← onboarding quick-start
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| CAT-GS   | Google Sheets only | $19 |
| CAT-PDF  | Printable PDF only | $19 |
| CAT-BUNDLE | Sheets + PDF + Quick-Start | **$29** |
| CAT-PLUS | Bundle + plate-costing add-on | $39 |
| CAT-PRO  | Commercial / multi-location license | $99 |

- **Steady all-year demand** with peaks before wedding season (spring) and the
  Q4 holiday events season. Priced above consumer planners — this saves an
  under-quoting caterer real money on a single event.
- Use all 10 photos + a walkthrough video. Lead photo = the feature-forward hero;
  the **plate-costing engine** and **quote = full P&L** are your strongest
  differentiators — most listings are a plain price list.
- Cross-sell the **Recipe Costing**, **Restaurant** and **Personal / Private Chef**
  products — same buyer, natural "food business" bundle.

---

## F. Maintenance

- Edit the `PLATE`, `PACKAGES`, `EVENTS`, `STAFFING`, `RENTALS`, `BOOKINGS`,
  `INVENTORY`, `WASTE`, `ORDERING`, `CASHDEP`, `CLIENTS` constants in
  `build_xlsx.py`; every KPI + the Catering Score recompute. Add a package → add a
  `PACKAGES` row (margins follow); add an event → add an `EVENTS` row (its P&L
  follows).
- Keep `build_marketing.py`'s KPIs in sync with the workbook.

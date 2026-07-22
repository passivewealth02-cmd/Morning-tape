# Personal & Private Chef Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/private-chef-command-center/build
python3 build_xlsx.py      # -> ../Private_Chef_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not financial/food-safety
   advice" note.
2. **Event Pricing** costs the flagship Private Dinner to **FOOD / GUEST $22.00**
   (named `FoodPerGuest`); at **$90/guest** for **6 guests** that's an event price of
   **$540**, less **$132** food and **$30** travel = **YOUR TAKE-HOME $378**, ÷ **7**
   hours = **YOUR REAL HOURLY RATE $54.00**. Food cost **24%**, margin **70%**.
3. **Clients** rolls up 5 clients to **MONTHLY REVENUE $5,950** (named `MonthlyRev`),
   **13** events/month, with the top client (`Anderson`) surfaced by `INDEX/MATCH`.
4. **Income & Expenses** takes `MonthlyRev` in and subtracts $2,610 of costs for a
   **MONTHLY PROFIT of $3,340**.
5. **Dashboard** fills 12 KPI cards + a Chef Health table + a revenue-by-month chart.
   **Chef Score 90%** (waste-low is the honest weak dimension).
6. No broken cells; custom tables (Event Pricing, Service Menu, Clients, etc.) start
   in column B.

> Note: uses `INDEX`, `MATCH`, `COUNTIF`, `COUNTA`, `AVERAGE`, `MAX`, `MIN`,
> `IFERROR` — opens in Google Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Private_Chef_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter.
Twelve pages: event quote, service menu, dish cost card, client roster, prep list,
shopping list, kitchen kit, event run sheet, mileage & travel, income & expenses,
monthly summary, and an event-day checklist.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with the price-breakdown & revenue
charts), everything-inside (14 tabs), the clients→revenue roster, the pay-yourself
pricing engine, the pricing engine (event pricing + service menu), and the **12-page
printables showcase**. Images 3–5 each show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "basic price list vs Command
Center", 09 run-your-chef-business in 4 steps, 10 what's-included / who-it's-for /
works-with. Ten images — fills all 10 Etsy slots. All headline numbers ($22/guest
food · $90/guest price · 24% food cost · 70% margin · $54.00/hr · $5,950 revenue · 5
clients · 13 events · $3,340 profit · 90% score) are verified against the workbook.

---

## D. Etsy delivery package

```
Private_Chef_Command_Center.xlsx   ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
Private_Chef_Printables.pdf         ← 12-page print-ready pack
START_HERE.pdf                      ← onboarding quick-start
```

---

## E. Pricing  (all SKUs are the same digital files — a product, never a service)

| SKU | What the buyer downloads | Price |
| --- | ------ | ----- |
| PCC-GS   | The Google Sheets / Excel file only | $19 |
| PCC-PDF  | The printable PDF only | $19 |
| PCC-BUNDLE | The spreadsheet + the printable PDF | **$29** |
| PCC-COMM | The same files + a commercial-use file license | $49 |

> Etsy Services-policy note: keep every listing a plain digital download. Do not
> advertise setup help, custom work, consultations, "done-for-you" builds, or "free
> updates / lifetime access" — offering a service (rather than a finished file) is
> what gets a listing removed and earns a strike. A commercial-use license is a
> permission term attached to the *file* and is allowed; doing the work *for* the
> buyer is not.

- **Steady all-year demand** with peaks around the holidays and wedding season.
  Priced above consumer planners — one under-priced dinner is real money.
- Use all 10 photos + a walkthrough video that shows the *file*. Lead photo = the
  feature-forward hero; the **pay-yourself pricing engine** and the **clients→revenue
  roster** are your strongest differentiators — most listings are just a price list.
- Cross-sell the **Catering**, **Meal Prep** and **Recipe Costing** products — same
  buyer, natural "food business" bundle.

---

## F. Maintenance

- Edit the `EVENT`, `GUESTS`, `PRICE_PER_GUEST`, `TRAVEL`, `HOURS`, `MENU`, `DISHES`,
  `CLIENTS`, `BOOKINGS`, `GROCERIES`, `KIT`, `MILEAGE`, `WASTE`, `LEDGER`, `MONTHS`
  constants and the `TARGET_FC`, `MARGIN_GOAL`, `CLIENT_GOAL`, `PROFIT_GOAL`,
  `WASTE_LIMIT` targets in `build_xlsx.py`; every KPI + the Chef Score recompute.
  Everything is cross-linked — change a course cost or a client and the take-home,
  the hourly rate and the score all follow.
- Keep `build_marketing.py`'s KPIs in sync with the workbook.

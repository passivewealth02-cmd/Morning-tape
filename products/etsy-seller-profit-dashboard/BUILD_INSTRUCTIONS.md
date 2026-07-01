# Etsy Seller Profit Dashboard™ — Build Instructions

---

## A. Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/etsy-seller-profit-dashboard/build
python3 build_xlsx.py      # -> ../Etsy_Seller_Profit_Dashboard.xlsx  (14-tab system + Welcome)
```

The build also prints a SUMMARY (revenue, fees, net profit, best/worst
product, etc.) computed from the sample orders — used to keep the marketing
numbers truthful to the workbook.

### Verifying
1. **Welcome** shows the intro + start-here steps.
2. **Settings** holds the 2026 Etsy fee rates (editable).
3. **Order Tracker** auto-calculates Etsy Fees & Net Profit per order; mark an
   order Refunded and its profit drops to $0 and the row flags red.
4. **Product Calc / Library** compute per-listing margin & rank products by
   profit (pulled from orders via `SUMIF`).
5. **Seller Dashboard** fills 10 KPIs (incl. Best Seller / Needs Work via
   `INDEX/MATCH`) + 4 charts. **Fees, Ads, Cash Flow, Tax Prep, Strategy** all
   compute from the same order data. No broken cells (blank-safe totals).

> Note: `MINIFS`/`INDEX`/`MATCH` — open in Excel 2019/365 or Google Sheets.

---

## B. Google Sheets build

Follow `GOOGLE_SHEETS.md`: build **Settings** first (fee rates + lists), define
the cross-sheet named ranges, then Orders → everything else.

---

## C. Marketing images

```bash
cd build && python3 build_marketing.py     # -> ../marketing/01..06.png
```

Six 2000×2000 PNGs, rendered as dense "finance-app screenshots" (sidebar of
all 14 tabs + the real computed numbers + fully populated tables/charts): hero,
everything-inside (14-tab showcase), executive dashboard, the Order Tracker
profit engine, fee breakdown + product ranking, and mobile.

---

## D. Etsy delivery package

```
Etsy_Seller_Profit_Dashboard.xlsx   ← Excel master (14-tab system + Welcome)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
START_HERE.pdf                      ← onboarding quick-start
THANK_YOU.pdf                       ← brand thank-you card
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| ESP-EX     | Excel only | $19 |
| ESP-GS     | Google Sheets only | $19 |
| ESP-BUNDLE | Excel + Google Sheets + Quick-Start PDF | **$27** |
| ESP-PLUS   | Bundle + "Etsy fees & pricing" mini-course | $39 |
| ESP-PRO    | Shop-owner / creator resale-with-attribution license | $79 |

- **Evergreen bestseller category** — every Etsy seller eventually asks "how
  much am I actually making?" Steady demand all year; bumps in **January**
  (new-year shop goals) and **Q4** (holiday selling season).
- Meta-advantage: you're selling a profit tool *to Etsy sellers, on Etsy* —
  screenshots that show real fee math convert hard.
- Run a launch sale (~40% off) + bundle with any other planner you sell.

---

## F. Maintenance

- Fee rates live in Settings — update them each year (Etsy changes fees).
- Edit `PRODUCTS` / `ORDER_DIST` in `build_xlsx.py` to reshape the sample.
- New sheet: add a builder and slot it into the `order` list in `main()`.

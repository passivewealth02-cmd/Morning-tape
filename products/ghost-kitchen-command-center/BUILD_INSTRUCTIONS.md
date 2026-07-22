# Ghost Kitchen Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/ghost-kitchen-command-center/build
python3 build_xlsx.py      # -> ../Ghost_Kitchen_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not accounting advice" note.
2. **Item Margin** starts from the Signature Burrito app price **$13.95**, subtracts
   the 25% commission (−$3.49), food (−$3.80) & packaging (−$0.65) to a true net of
   **$6.01** (43% of app price), named `SigItemNet`.
3. **Menu & Margins** shows app price, food, packaging & the true net margin $ / %
   on 8 items — using the `Commission` setting. Avg net margin **44%**, avg food
   cost **26%**; top item (by net $) = Birria Tacos.
4. **Platform P&L** rolls up 4 platforms: weekly orders **800**, gross **$17,360**,
   blended commission **24%**, net payout **$13,117**, avg order **$21.70**,
   direct-order share **15%**.
5. **Dashboard** fills 12 KPI cards + a Kitchen Health table + a net-payout-by-
   platform chart. **Kitchen Score 90%** (commission-in-control is the honest weak
   dimension — the apps take ~24%).
6. No broken cells; custom tables (Item Margin, Menu & Margins, Platform P&L, etc.)
   start in column B.

> Note: uses `INDEX`, `MATCH`, `COUNTIF`, `COUNTA`, `AVERAGE`, `MAX`, `MIN`,
> `IFERROR` — opens in Google Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Ghost_Kitchen_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter.
Twelve pages: item-margin card, menu & margins, platform P&L, virtual brands,
packaging sheet, order-volume log, prep list, inventory & par, waste log, ordering
sheet, payouts sheet, and a promotions sheet.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with commission & net-payout charts),
everything-inside (14 tabs), the menu net margins, the platform P&L, the
item-margin engine, and the **12-page printables showcase**. Images 3–5 each show a
different tab.

**Four detailed images**: 07 feature spotlights, 08 "basic menu vs Command Center",
09 beat-the-apps in 4 steps, 10 what's-included / who-it's-for / guarantee. Ten
images — fills all 10 Etsy slots. All headline numbers (8 items · $13.95 app price
· 24% commission · 44% net margin · $17,360 revenue · $13,117 net payout · 90%
score) are verified against the workbook.

---

## D. Etsy delivery package

```
Ghost_Kitchen_Command_Center.xlsx  ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
Ghost_Kitchen_Printables.pdf        ← 12-page print-ready pack
START_HERE.pdf                      ← onboarding quick-start
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| GKC-GS   | Google Sheets only | $19 |
| GKC-PDF  | Printable PDF only | $19 |
| GKC-BUNDLE | Sheets + PDF + Quick-Start | **$29** |
| GKC-PLUS | Bundle + margin add-on | $39 |
| GKC-PRO  | Commercial / multi-location license | $99 |

- **Steady all-year demand** — delivery is a year-round business with a January
  new-year planning peak. Priced above consumer planners — this saves a kitchen
  real money on every order the apps touch.
- Use all 10 photos + a walkthrough video. Lead photo = the feature-forward hero;
  the **item-margin-after-commission** engine and **platform P&L** are your
  strongest differentiators — most listings are a plain menu.
- Cross-sell the **Restaurant**, **Food Cost & Inventory** and **Recipe Costing**
  products — same buyer, natural "food business" bundle.

---

## F. Maintenance

- Edit the `SIG_ITEM`, `MENU`, `PLATFORMS`, `BRANDS`, `PACKAGING`, `ORDERVOL`,
  `INVENTORY`, `WASTE`, `ORDERING`, `PAYOUTS`, `PROMOS` constants in
  `build_xlsx.py`; every KPI + the Kitchen Score recompute. Change the default
  `Commission` in Settings and every item's net margin follows.
- Keep `build_marketing.py`'s KPIs in sync with the workbook.

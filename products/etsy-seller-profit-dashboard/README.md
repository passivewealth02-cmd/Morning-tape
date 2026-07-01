# Etsy Seller Profit Dashboard™ — The Ultimate Etsy Finance & Profit System (2026)

> Not a tracker — a **mini Etsy CFO system** in Excel & Google Sheets. It turns messy shop data into clean profit intelligence and answers the one question every seller has: *"How much am I actually making after fees?"*

| | |
| - | - |
| **Product** | Etsy Seller Profit Dashboard™ |
| **Target** | Etsy beginners (0–100 sales/mo) · side hustlers · digital / printable / POD / handmade sellers · anyone scaling to full-time |
| **Angle** | See real **profit**, not just revenue — after Etsy fees, ads, refunds & taxes. |
| **Formats** | Excel `.xlsx` (14-tab system + Welcome) + Google Sheets edition |
| **Pricing** | $19 single · **$27 bundle** · $39 with mini-course · $79 shop-owner/creator license |

---

## Contents

```
products/etsy-seller-profit-dashboard/
├── README.md
├── Etsy_Seller_Profit_Dashboard.xlsx   ← Excel master (14-tab system + Welcome)
├── GOOGLE_SHEETS.md
├── BUILD_INSTRUCTIONS.md
├── ETSY_LISTING.md
└── build/
    ├── build_xlsx.py
    └── build_marketing.py
```

---

## The 14-sheet system

| # | Sheet | # | Sheet |
| - | ----- | - | ----- |
| 1 | Seller Executive Dashboard | 8 | Monthly Financial Snapshot |
| 2 | Order Tracker (the engine) | 9 | Customer & Order Insights |
| 3 | Product Profit Calculator | 10 | Goal Tracker |
| 4 | Etsy Fees Breakdown | 11 | Cash Flow Dashboard |
| 5 | Ads Performance Tracker | 12 | Tax Preparation Sheet |
| 6 | Expense Tracker | 13 | Strategy Analyzer |
| 7 | Product Library | 14 | Settings |

*(+ a Welcome / Start-Here tab — 15 tabs total.)*

---

## The automation that sells it

**The Order Tracker is the engine.** Type in a sale (product, qty, price) and
it auto-calculates — using your fee rates from Settings — the:

```
Etsy Fees = Listing×Qty + Txn%×(Sale+Ship) + (Proc%×(Sale+Ship)+ProcFixed)
            + Offsite%×(Sale+Ship)   [if the order came from an offsite ad]
Net Profit = Sale + Shipping − Etsy Fees − COGS      (refunds auto → $0)
```

Everything else flows from there:

| Metric | Source |
| ------ | ------ |
| Net Profit (after everything) | `SUM(OrdProfit) − Etsy Ads − Expenses` |
| Profit Margin | Net Profit ÷ Revenue |
| Etsy Fees Paid | `FeesTotal` (broken down by type) |
| Best Seller / Needs Work | `INDEX/MATCH` on the Product Library |
| Cost per game… er, order | AOV, cost of selling, retention |
| ROAS | `Ad Revenue ÷ Ad Spend` |
| Tax reserve | `Net Profit × your set-aside %` |

**48 named ranges**, blank-safe formulas, clean profit logic (no double
counting), and a **Strategy Analyzer** that names your best/worst listings.

---

## SaaS-style finance design

- Clean **KPI cards** + an executive dashboard (QuickBooks-lite, but simpler)
- Gold-topped cards, soft shadows, minimal clutter
- Refunds flag red, high ROAS flags mint, over-budget flags red
- 2026 Etsy fee rates pre-filled (fully editable)
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> Tracking tool only — not tax or accounting advice. Confirm your exact fee
> rates in your Etsy Payment account.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py        # -> ../Etsy_Seller_Profit_Dashboard.xlsx
python3 build_marketing.py               # -> ../marketing/01..06.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

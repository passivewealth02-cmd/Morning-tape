# Amazon FBA & Online Seller Profit Command Center™ — The Seller's Operating System

> Not a fee calculator — a **complete price-it, stock-it, scale-it system**.
> One premium **Google Sheets + printable PDF** command center for product sellers: a
> true-profit engine (sale price − referral − FBA − storage − COGS − inbound → net per
> unit, margin and ROI), a product catalog, a fee breakdown, inventory & reorder, sales,
> PPC & ACoS, returns, suppliers, reviews, expenses and a monthly summary — everything
> cross-linked and live.

| | |
| - | - |
| **Product** | Amazon FBA & Online Seller Profit Command Center™ |
| **Target** | Amazon FBA sellers · private-label brand owners · Shopify & Walmart sellers · new sellers costing their first SKU · sellers scaling past six figures |
| **Angle** | Know your real profit per unit — after every fee Amazon takes. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $22 single file · **$32 bundle** (Sheets + PDF) · $55 with a commercial-use file license — all the *same digital files*, never a service |

---

## Contents

```
products/fba-seller-command-center/
├── README.md
├── FBA_Seller_Command_Center.xlsx    ← Google Sheets / Excel master (14 tabs)
├── FBA_Seller_Printables.pdf         ← 12-page print-ready pack (US Letter)
├── GOOGLE_SHEETS.md
├── BUILD_INSTRUCTIONS.md
├── ETSY_LISTING.md
└── build/
    ├── build_xlsx.py
    ├── build_pdf.py
    ├── build_marketing.py
    └── build_marketing_detail.py
```

---

## The 14-tab system

| # | Tab | # | Tab |
| - | --- | - | --- |
| 1 | Start Here | 8 | PPC & ACoS |
| 2 | Dashboard | 9 | Returns |
| 3 | Profit Calculator | 10 | Suppliers & POs |
| 4 | Product Catalog | 11 | Reviews |
| 5 | Fee Breakdown | 12 | Expenses |
| 6 | Inventory | 13 | Monthly Summary |
| 7 | Sales | 14 | Settings |

## The 12 printable PDF pages

Profit Per Unit · Product Catalog · Fee Worksheet · Inventory & Reorder · PPC & ACoS Log
· Sales Log · Supplier & PO Log · Returns Log · Reviews Tracker · Product Research ·
Monthly Summary · Launch Checklist.

---

## Signature automation — the true-profit engine

Everything connects. A sale price becomes a net per unit only after every fee and cost
comes out — and that net is what decides whether a product is worth selling:

```
Amazon fees  = (price × referral rate) + FBA fee + storage
Landed cost  = cost of goods + inbound shipping
Net per unit = price − Amazon fees − landed cost
Net margin   = net per unit ÷ price
ROI          = net per unit ÷ landed cost
ACoS         = ad spend ÷ ad-attributed sales
Days of cover = units on hand ÷ units sold per day
```

### The 12 dashboard KPIs
Sale Price · Amazon Fees · Landed Cost · Net/Unit · Net Margin · ROI · Units/Month ·
Monthly Revenue · Monthly Profit · ACoS · Days of Cover · Seller Score.
The **Seller Score** blends margin, ROI, ACoS, inventory cover, catalog depth and review
velocity into one 0–100% number.

**Verified sample brand** (Northport Goods, owner Sam, flagship Soy Candle 8oz): sale
price **$29.99** · Amazon fees **$10.25** · landed cost **$8.70** · net/unit **$11.04** ·
net margin **36.8%** · ROI **126.9%** · units **420**/mo · revenue **$12,596** · monthly
profit **$4,637** · ACoS **20.0%** · days of cover **60** · **Seller Score 90%** (the
honest weak spot: 4 reviews per SKU against a goal of 10).

---

## Premium seller design

- A **true-profit engine** that strips out every fee, not just the referral
- **Net per unit, margin AND ROI** — the three numbers that decide a product
- A whole **catalog costed** the same way, with losers flagged red
- **Days-of-cover** inventory and per-SKU reorder points
- **PPC & ACoS** plus **TACoS** against total revenue, returns, suppliers and reviews
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A business & organizing tool, not financial, tax or accounting advice.** Amazon's
> fee schedule changes — check your own Seller Central fee preview and update Settings.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../FBA_Seller_Command_Center.xlsx
python3 build_pdf.py                         # -> ../FBA_Seller_Printables.pdf (+ page PNGs)
python3 build_marketing.py                  # -> ../marketing/01..06.png
python3 build_marketing_detail.py           # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

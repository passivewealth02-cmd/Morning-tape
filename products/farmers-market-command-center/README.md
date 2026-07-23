# Farmers Market Vendor Command Center™ — The Market-Stall Operating System

> Not a price list — a **complete cost-it, sell-it, did-it-pay system**.
> One premium **Google Sheets + printable PDF** command center for a market stall: a
> product price & margin list, a per-market booth P&L engine (enter what you sold and
> see if the day actually paid after your cost of goods and booth costs), a bake/prep
> plan, ingredient & packaging costs, a markets log that becomes monthly sales, your
> regulars & CSA subscribers, waste, income & expenses and a monthly summary —
> everything cross-linked and live.

| | |
| - | - |
| **Product** | Farmers Market Vendor Command Center™ |
| **Target** | Farmers-market bakers · jam, honey & preserve makers · produce & flower growers · cottage-food & craft sellers · CSA & farm-stand sellers · anyone with a market stall |
| **Angle** | Cost every product, and know if the market day paid. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $19 single file · **$29 bundle** (Sheets + PDF) · $49 with a commercial-use file license — all the *same digital files*, never a service |

---

## Contents

```
products/farmers-market-command-center/
├── README.md
├── Farmers_Market_Command_Center.xlsx  ← Google Sheets / Excel master (14 tabs)
├── Farmers_Market_Printables.pdf       ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 8 | Booth Costs |
| 2 | Dashboard | 9 | Packaging |
| 3 | Market Day | 10 | Customers |
| 4 | Products | 11 | Waste Log |
| 5 | Ingredient Costs | 12 | Income & Expenses |
| 6 | Bake Plan | 13 | Monthly Summary |
| 7 | Markets | 14 | Settings |

## The 12 printable PDF pages

Booth P&L · Product Price List · Bake Plan · Ingredient List · Markets Log · Booth
Costs · Packaging · Customer List · Waste Log · Income & Expenses · Monthly Summary ·
Market Day Checklist.

---

## Signature automation — did the market day pay?

Everything connects. Your product prices flow into the market-day sales, your costs
come out, and the booth P&L tells you if it was worth going:

```
Market sales     = Σ (units sold × product price)
Cost of goods    = Σ (units sold × product cost)
Booth net        = Market sales − COGS − booth costs (stall, fuel, help)
Net per hour     = Booth net ÷ hours worked
Monthly sales    = Σ market-day sales
Product margin % = (Price − Cost) ÷ Price
```

### The 12 dashboard KPIs
Units Sold · Market Sales · COGS % · Booth Net · Net/Hour · Top Seller · Avg Basket ·
Monthly Sales · Markets/Month · Monthly Profit · Waste % · Vendor Score. The
**Vendor Score** blends COGS-on-target, margin-healthy, products-priced,
sales-vs-goal, profitable and waste-low into one 0–100% number.

**Verified sample stall** (Harvest Lane Market Co., owner Sage): **90** units · market
sales **$768** · COGS **24%** · booth net **$438** · net/hour **$54.75** · top seller
**Sourdough** ($176) · avg basket **$12.00** · monthly sales **$3,000** · **4**
markets · monthly profit **$1,595** · waste **2.4%** · **Vendor Score 90%**.

---

## Premium market-vendor design

- A **per-market booth P&L** engine (did this day pay, and what per hour?)
- **Product margins** on every item you sell
- A **bake / prep plan** and **ingredient & packaging** costs
- A **markets log** that rolls each day up into **monthly sales**
- **Regulars & CSA**, **waste**, **income & expenses** and a **monthly summary**
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A business tool, not financial, legal or food-safety advice.** Follow your local
> cottage-food and market rules and confirm figures with your own advisors.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Farmers_Market_Command_Center.xlsx
python3 build_pdf.py                         # -> ../Farmers_Market_Printables.pdf (+ page PNGs)
python3 build_marketing.py                  # -> ../marketing/01..06.png
python3 build_marketing_detail.py           # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

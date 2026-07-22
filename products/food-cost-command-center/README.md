# Food Cost & Inventory Command Center™ — The Food-Cost Control System

> Not a count sheet — a **complete count-it, cost-it, control-it system**.
> One premium **Google Sheets + printable PDF** command center for food-cost
> control: a period food-cost engine (Beginning + Purchases − Ending ÷ Sales),
> inventory valuation to the dollar, a purchases & sales log, usage-vs-theoretical
> variance, par & ordering, vendors, a price tracker, menu costing, waste and a
> category breakdown — everything cross-linked and live.

| | |
| - | - |
| **Product** | Food Cost & Inventory Command Center™ |
| **Target** | Restaurants & cafés · bars & pubs · food trucks & caterers · bakeries & ghost kitchens · chefs & kitchen managers · anyone who tracks food cost |
| **Angle** | Count it, cost it, control it — know your food cost to the point. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $19 single · **$29 bundle** (Sheets + PDF) · $39 with the cost add-on · $99 commercial / multi-location license |

---

## Contents

```
products/food-cost-command-center/
├── README.md
├── Food_Cost_Command_Center.xlsx     ← Google Sheets / Excel master (14 tabs)
├── Food_Cost_Printables.pdf          ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 8 | Par & Ordering |
| 2 | Dashboard | 9 | Vendors |
| 3 | Food Cost Calc | 10 | Price Tracker |
| 4 | Inventory Count | 11 | Menu Costing |
| 5 | Purchases Log | 12 | Waste Log |
| 6 | Sales Log | 13 | Categories |
| 7 | Usage & Variance | 14 | Settings |

## The 12 printable PDF pages

Inventory Count Sheet · Food Cost Worksheet · Purchases Log · Sales Log · Usage &
Variance · Par & Order Guide · Vendor List · Price Comparison · Menu Costing ·
Waste Log · Category Breakdown · Weekly Count Checklist.

---

## Signature automation — food cost to the point

Everything connects. Your counts value your ending inventory, your invoices total
your purchases, and your sales feed the denominator — so the food-cost % is always
live:

```
Inventory value = Σ (count × unit cost)
Food used (COGS) = Beginning + Purchases − Ending inventory
Food cost %     = Food used ÷ Sales
Variance %      = (Actual usage − Theoretical usage) ÷ Theoretical
Inventory turns = Food used ÷ Avg inventory
```

### The 12 dashboard KPIs
Food Cost % · Inventory Value · Purchases · Food Used · Sales · Top Category ·
Variance · Items Tracked · To Order · Inv Turns · Vendors · Inventory Score. The
**Inventory Score** blends food-cost-on-target, low-variance, inventory-counted,
pars-set, turns-healthy and gross-margin into one 0–100% number.

**Verified sample kitchen** (The Harvest Table, owner Sam): food cost **30.0%** ·
inventory value **$11,800** · purchases **$18,600** · food used **$19,200** · sales
**$64,000** · top category **Meat & seafood** ($7,200) · variance **2.9%** · **14**
items · to order **$2,705** · turns **1.6** · **6** vendors · **Inventory Score 90%**.

---

## Premium food-cost design

- A **period food-cost** engine (Beginning + Purchases − Ending ÷ Sales)
- **Inventory valuation** to the dollar (count × unit cost)
- **Usage vs theoretical** variance to catch shrinkage
- **Par & ordering**, **vendors**, a **price tracker** & **menu costing**
- **Waste** & a **category breakdown** of where the food dollar goes
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A business tool, not financial or accounting advice.** Confirm figures with
> your own books.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Food_Cost_Command_Center.xlsx
python3 build_pdf.py                        # -> ../Food_Cost_Printables.pdf (+ page PNGs)
python3 build_marketing.py                 # -> ../marketing/01..06.png
python3 build_marketing_detail.py          # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

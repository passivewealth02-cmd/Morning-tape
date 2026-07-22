# Bakery Command Center™ — The Complete Bakery Business System

> Not a price list — a **complete cost-batch, price-retail-and-wholesale system**.
> One premium **Google Sheets + printable PDF** command center for a bakery: a
> recipe cost-per-batch engine, a product list with retail & wholesale margins, a
> pre-orders board, wholesale accounts, a production plan, inventory & par, a
> waste log, a sales log, ordering, cash & deposits and market days.

| | |
| - | - |
| **Product** | Bakery Command Center™ |
| **Target** | Bakeries & bakery cafés · home & cottage bakers · bread & pastry shops · cake & cookie makers · wholesale & farmers-market bakers · anyone pricing baked goods |
| **Angle** | Cost every batch, price retail & wholesale, and bake more profit. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $19 single · **$29 bundle** (Sheets + PDF) · $39 with the recipe-costing add-on · $99 commercial / multi-location license |

---

## Contents

```
products/bakery-command-center/
├── README.md
├── Bakery_Command_Center.xlsx        ← Google Sheets / Excel master (14 tabs)
├── Bakery_Printables.pdf             ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 8 | Inventory & Par |
| 2 | Dashboard | 9 | Waste Log |
| 3 | Recipe Costing | 10 | Sales Log |
| 4 | Product List | 11 | Ordering |
| 5 | Pre-Orders | 12 | Cash & Deposits |
| 6 | Wholesale | 13 | Market Days |
| 7 | Production Plan | 14 | Settings |

## The 12 printable PDF pages

Recipe Cost Card · Product Price List · Pre-Order Form · Wholesale Order Sheet ·
Production Plan · Inventory & Par · Waste & Day-Old Log · Sales Log · Ordering
Sheet · Cash & Deposits · Market Day Sheet · Bake-Day Checklist.

---

## Signature automation — cost by the batch, price both channels

The recipe engine costs a recipe by the batch, then divides by yield for a true
cost per unit; each product then carries both a **retail** and a **wholesale**
price with its own margin and food-cost %:

```
Cost per unit       = Batch cost ÷ Batch yield
Retail margin       = Retail price − Cost per unit
Food-cost %         = Cost per unit ÷ Retail price
Weekly revenue      = Retail revenue + Wholesale revenue
```

### The 12 dashboard KPIs
Products · Avg Unit Cost · Avg Retail · Avg Margin · Food Cost · Top Seller ·
Weekly Revenue · Weekly Units · Pre-Orders · Wholesale Rev · Waste % · Bakery
Score. The **Bakery Score** blends food-cost-on-target, margin-per-unit,
fully-costed, pre-orders-vs-goal, low-waste and gross-margin into one 0–100%
number.

**Verified sample bakery** (Rise & Crumb, owner Nora): **8** products · avg unit
cost **$1.13** · avg retail **$5.19** · avg margin **$4.06** · food cost **22%** ·
top seller **Butter Croissant** ($1,275/wk) · weekly revenue **$8,677** (retail
$6,495 + wholesale $2,182) · **1,480** units · **6** pre-orders · waste **4.5%** ·
**Bakery Score 88%**.

---

## Premium bakery-software design

- A **recipe cost-per-batch** engine (batch ÷ yield = true cost per unit)
- A **product list** with retail *and* wholesale prices, margins & food-cost %
- A **pre-orders** board and a **wholesale accounts** tracker
- A **production plan**, **inventory & par**, **waste**, **sales** & **market days**
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A business tool, not financial or accounting advice.** Confirm figures with
> your own books.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Bakery_Command_Center.xlsx
python3 build_pdf.py                        # -> ../Bakery_Printables.pdf (+ page PNGs)
python3 build_marketing.py                 # -> ../marketing/01..06.png
python3 build_marketing_detail.py          # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

# Restaurant Command Center™ — The Ultimate Restaurant Operations, Cost & Team Management System

> Not a spreadsheet — a **complete Restaurant Operating System**. One premium
> Excel & Google Sheets command center for sales, food & labor cost, inventory,
> P&L, scheduling, reservations, marketing, reviews, compliance and your team —
> with restaurant-grade automation that protects your margins.

| | |
| - | - |
| **Product** | Restaurant Command Center™ |
| **Target** | Restaurant & café owners · bar / pub operators · food trucks & ghost kitchens · caterers · franchisees & multi-unit GMs · new-restaurant openers |
| **Angle** | Control prime cost, cut waste and run a tighter, more profitable operation. |
| **Formats** | Excel `.xlsx` (22-tab system + Welcome) + Google Sheets edition |
| **Pricing** | $29 single · **$49 bundle** · $79 with SOP & manager toolkit · $149 multi-unit / consultant license |

---

## Contents

```
products/restaurant-command-center/
├── README.md
├── Restaurant_Command_Center.xlsx   ← Excel master (22-tab system + Welcome)
├── GOOGLE_SHEETS.md
├── BUILD_INSTRUCTIONS.md
├── ETSY_LISTING.md
└── build/
    ├── build_xlsx.py
    └── build_marketing.py
```

---

## The 22-sheet system

| # | Sheet | # | Sheet |
| - | ----- | - | ----- |
| 1 | Restaurant Dashboard | 12 | Cash, Tips & Deposits |
| 2 | Restaurant Profile | 13 | Reservations & Events |
| 3 | Menu & Recipe Costing | 14 | Marketing & Promos |
| 4 | Inventory Master | 15 | Reviews & Reputation |
| 5 | Par Levels & Order Guide | 16 | Health & Safety Compliance |
| 6 | Suppliers & Ordering | 17 | Cleaning & Maintenance |
| 7 | Sales Tracker | 18 | Waste & Loss Tracker |
| 8 | Labor & Scheduling | 19 | Vendor Payments (AP) |
| 9 | Staff Roster | 20 | Training & Certifications |
| 10 | P&L / Prime Cost | 21 | Operations Analytics |
| 11 | Expenses & Overhead | 22 | Settings |

*(+ a Welcome / Start-Here tab — 23 tabs total.)*

---

## Signature automation (12 dashboard KPIs)

| Metric | How it's computed |
| ------ | ----------------- |
| Monthly Revenue | `=RevenueTotal` (from daily sales) |
| Covers | `=CoversTotal` |
| Avg Check | `=RevenueTotal/CoversTotal` |
| Food Cost % | `=FoodCostTotal/RevenueTotal` |
| Labor Cost % | `=LaborTotal/RevenueTotal` |
| Prime Cost % | `=(FoodCostTotal+LaborTotal)/RevenueTotal` |
| Net Profit | `=RevenueTotal-FoodCostTotal-LaborTotal-OpexTotal` |
| Profit Margin | `=NetProfit/RevenueTotal` |
| Inventory Value | `=SUM(OnHand×UnitCost)` |
| Low-Stock Items | `=COUNTIF(InvStatus,"Low")` |
| Avg Rating | `=AVERAGE(platform ratings)` |
| Operations Health Score | `=AVERAGE(HealthRange)` |

Menu & Recipe Costing computes **food cost % & contribution margin per dish**;
inventory auto-flags **low stock** and values every item; the P&L rolls sales,
food, labor and overhead into **prime cost, gross profit & net margin**; and an
**Operations Health Score** blends food/labor/prime-cost control, profit margin,
guest ratings and inventory levels.

**Verified sample restaurant** (Olive & Ember, an 84-seat modern bistro):
Revenue **$131,510** · Avg Check **$30.34** · Food **32.1%** · Labor **32.8%** ·
Prime **64.9%** · Net Profit **$19,526** (14.8%) · Health Score **91%**.

---

## Premium restaurant-software design

- Two-row **gold-divider headers** on every tab
- Gold-topped white **KPI cards** + a true operations dashboard (12 KPIs +
  cost-structure, weekly-sales, health & prime-cost charts)
- Low-stock flags red, in-stock mint; food-cost heat-map on the menu; payment
  status & compliance flags; nightly cash reconciliation
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

---

## Build & ship

```bash
cd build && python3 build_xlsx.py        # -> ../Restaurant_Command_Center.xlsx
python3 build_marketing.py               # -> ../marketing/01..06.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

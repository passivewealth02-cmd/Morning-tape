# Uber Eats Driver Command Center™ — The Ultimate Food-Delivery Earnings & Budget System

> Not a mileage app — a **complete food-delivery operating system**. One premium
> Excel & Google Sheets command center for shifts, earnings, mileage, gas,
> vehicle upkeep, business expenses, a household budget, taxes, savings and
> promos — everything that turns "cash today" into a **profitable delivery
> business**. Built for Uber Eats; works for DoorDash, Grubhub, Instacart or any gig.

| | |
| - | - |
| **Product** | Uber Eats Driver Command Center™ |
| **Target** | Full-time & part-time delivery drivers · Uber Eats & DoorDash drivers · Grubhub & Instacart shoppers · gig-economy multi-appers · anyone who deducts mileage |
| **Angle** | Know your REAL take-home — net $/hour, the mileage tax deduction & a budget that works. |
| **Formats** | Excel `.xlsx` (18-tab system + Welcome) + Google Sheets edition |
| **Pricing** | $15 single · **$22 bundle** · $29 with tax & promo playbook · $79 fleet / referral license |

---

## Contents

```
products/uber-eats-driver-command-center/
├── README.md
├── Uber_Eats_Driver_Command_Center.xlsx  ← Excel master (18-tab system + Welcome)
├── GOOGLE_SHEETS.md
├── BUILD_INSTRUCTIONS.md
├── ETSY_LISTING.md
└── build/
    ├── build_xlsx.py
    ├── build_marketing.py
    └── build_marketing_detail.py
```

---

## The 18-sheet system

| # | Sheet | # | Sheet |
| - | ----- | - | ----- |
| 1 | Driver Dashboard | 10 | Tax Center |
| 2 | Driver Profile | 11 | Savings & Goals |
| 3 | Delivery Log | 12 | Promos & Quests |
| 4 | Earnings Breakdown | 13 | Hotspots & Hours |
| 5 | Mileage Tracker | 14 | Ratings & Feedback |
| 6 | Fuel Log | 15 | Analytics Center |
| 7 | Vehicle & Maintenance | 16 | Weekly Planner |
| 8 | Business Expenses | 17 | Receipts Gallery |
| 9 | Monthly Budget | 18 | Settings |

*(+ a Welcome / Start-Here tab — 19 tabs total.)*

---

## Signature automation (12 dashboard KPIs)

| Metric | How it's computed |
| ------ | ----------------- |
| Net Earnings | `=GrossEarn-ExpenseTotal` |
| Gross Earnings | `=SUM(ShiftEarn)` (base+tips+promo) |
| Tips | `=SUM(ShiftTips)` |
| Net $ / Hour | `=(GrossEarn-ExpenseTotal)/TotalHours` |
| Net $ / Mile | `=(GrossEarn-ExpenseTotal)/TotalMiles` |
| Miles Driven | `=SUM(ShiftMiles)` |
| Online Hours | `=SUM(ShiftHours)` |
| Deliveries | `=SUM(ShiftOrders)` |
| Expenses | `=ExpenseTotal` (gas auto-linked from the log) |
| Tax Deduction | `=TotalMiles*MileageRate` (IRS standard) |
| Savings | `=MIN(SavingsSaved/SavingsGoal,1)` |
| Driver Health Score | `=AVERAGE(HealthRange)` |

Every shift you log rolls into **live earnings, miles & gas**; the Expenses
sheet pulls gas straight from the log so your **cost-per-mile and true net**
stay honest; the Tax Center auto-picks **mileage vs actual** and banks a
quarterly set-aside; the Monthly Budget flows your **net income** into a plan;
and a **Driver Health Score** blends net earnings, $/hour, deliveries, savings,
consistency and your tax reserve.

**Verified sample driver** (Alex, a full-time Denver delivery driver, 16 shifts/mo):
Gross **$2,544** (Tips **$1,258** — 49% of pay) · Net **$1,526** · Net **$16.96/hr** ·
Net **$1.02/mi** · **1,502** miles · **321** deliveries · **90.0** hours ·
Expenses **$1,018** · Mileage deduction **$1,051** · Savings **65%** · Driver Health **87%**.

---

## Premium driver-software design

- Two-row **gold-divider headers** on every tab
- Gold-topped white **KPI cards** + a true driver dashboard (12 KPIs +
  net-earnings-trend, earnings-mix, best-shifts & expense charts)
- App-coded delivery log with earnings data-bars; demand & promo statuses
  color-code; a Tax Center deduction checklist so nothing gets left on the table
- Image-placeholder **Receipts & Vehicle Gallery** (Insert ▸ Picture-in-cell or
  `=IMAGE()`) for painless tax-time records
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Uber_Eats_Driver_Command_Center.xlsx
python3 build_marketing.py                 # -> ../marketing/01..06.png
python3 build_marketing_detail.py          # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

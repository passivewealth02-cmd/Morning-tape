# Instacart Shopper Command Center™ — The Ultimate Grocery-Delivery Earnings & Budget System

> Not a mileage app — a **complete grocery-shopper operating system**. One premium
> Excel & Google Sheets command center for shifts, batches, earnings, mileage,
> gas, vehicle upkeep, business expenses, a household budget, taxes, savings and
> boosts — everything that turns "cash today" into a **profitable shopping
> business**. Built for Instacart; works for Shipt, Amazon Flex or any grocery gig.

| | |
| - | - |
| **Product** | Instacart Shopper Command Center™ |
| **Target** | Full-time & part-time Instacart shoppers · Shipt shoppers · Amazon Flex & grocery-delivery drivers · gig-economy multi-appers · anyone who deducts mileage |
| **Angle** | Know your REAL take-home — net $/hour, $/batch, the tax deduction & a budget that works. |
| **Formats** | Excel `.xlsx` (18-tab system + Welcome) + Google Sheets edition |
| **Pricing** | $15 single · **$22 bundle** · $29 with tax & boost playbook · $79 fleet / referral license |

---

## Contents

```
products/instacart-shopper-command-center/
├── README.md
├── Instacart_Shopper_Command_Center.xlsx  ← Excel master (18-tab system + Welcome)
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
| 1 | Shopper Dashboard | 10 | Tax Center |
| 2 | Shopper Profile | 11 | Savings & Goals |
| 3 | Batch Log | 12 | Boosts & Bonuses |
| 4 | Earnings Breakdown | 13 | Best Stores & Hours |
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
| Gross Earnings | `=SUM(ShiftEarn)` (batch pay+tips+boost) |
| Tips | `=SUM(ShiftTips)` |
| Net $ / Hour | `=(GrossEarn-ExpenseTotal)/TotalHours` |
| Net $ / Mile | `=(GrossEarn-ExpenseTotal)/TotalMiles` |
| Miles Driven | `=SUM(ShiftMiles)` |
| Online Hours | `=SUM(ShiftHours)` |
| Batches | `=SUM(ShiftBatches)` |
| Expenses | `=ExpenseTotal` (gas auto-linked from the log) |
| Tax Deduction | `=TotalMiles*MileageRate` (IRS standard mileage) |
| Savings | `=MIN(SavingsSaved/SavingsGoal,1)` |
| Shopper Health Score | `=AVERAGE(HealthRange)` |

Every shift you log rolls into **live earnings, items, miles & gas**; the
Expenses sheet pulls gas straight from the log so your **cost-per-mile and true
net** stay honest; the Tax Center **compares mileage vs actual and auto-picks the
bigger deduction** (for a low-mileage shopper, that's often the actual method)
and banks a quarterly set-aside; the Monthly Budget flows your **net income**
into a plan; and a **Shopper Health Score** blends net earnings, $/hour, batches,
savings, consistency and your tax reserve.

**Verified sample shopper** (Riley, a full-time Portland shopper, 16 shifts/mo):
Gross **$2,472** (Tips **$1,317** — 53% of pay) · Net **$1,604** · Net
**$17.82/hr** · **$25.22/batch** · **98** batches / **2,440** items · **860**
miles · **90.0** hours · Expenses **$868** · Savings **70%** · Shopper Health **87%**.

---

## Premium shopper-software design

- Two-row **gold-divider headers** on every tab
- Gold-topped white **KPI cards** + a true shopper dashboard (12 KPIs +
  net-earnings-trend, earnings-mix, best-shifts & expense charts)
- App-coded batch log with earnings data-bars; demand & boost statuses
  color-code; a Tax Center that compares mileage vs actual so you claim the bigger one
- Image-placeholder **Receipts & Vehicle Gallery** (Insert ▸ Picture-in-cell or
  `=IMAGE()`) for painless tax-time records
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Instacart_Shopper_Command_Center.xlsx
python3 build_marketing.py                 # -> ../marketing/01..06.png
python3 build_marketing_detail.py          # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

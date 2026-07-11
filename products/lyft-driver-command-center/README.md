# Lyft Driver Command Center™ — The Ultimate Rideshare Earnings & Budget System

> Not a mileage app — a **complete rideshare-driver operating system**. One premium
> Excel & Google Sheets command center for shifts, earnings, mileage, fuel, vehicle
> upkeep, business expenses, a household budget, taxes, savings and bonuses —
> everything that turns "cash today" into a **profitable driving business**.
> Built for Lyft; works for Uber, delivery or any gig.

| | |
| - | - |
| **Product** | Lyft Driver Command Center™ |
| **Target** | Full-time & part-time rideshare drivers · Lyft & Uber drivers · delivery drivers (DoorDash, Uber Eats) · gig-economy side hustlers · anyone who deducts mileage |
| **Angle** | Know your REAL take-home — net $/hour, the mileage tax deduction & a budget that works. |
| **Formats** | Excel `.xlsx` (18-tab system + Welcome) + Google Sheets edition |
| **Pricing** | $15 single · **$22 bundle** · $29 with tax & bonus playbook · $79 fleet / referral license |

---

## Contents

```
products/lyft-driver-command-center/
├── README.md
├── Lyft_Driver_Command_Center.xlsx  ← Excel master (18-tab system + Welcome)
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
| 3 | Shift Log | 12 | Bonuses & Quests |
| 4 | Earnings Breakdown | 13 | Hot Zones & Hours |
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
| Gross Earnings | `=SUM(ShiftEarn)` (fares+tips+bonus) |
| Tips | `=SUM(ShiftTips)` |
| Net $ / Hour | `=(GrossEarn-ExpenseTotal)/TotalHours` |
| Net $ / Mile | `=(GrossEarn-ExpenseTotal)/TotalMiles` |
| Miles Driven | `=SUM(ShiftMiles)` |
| Online Hours | `=SUM(ShiftHours)` |
| Trips | `=SUM(ShiftTrips)` |
| Expenses | `=ExpenseTotal` (fuel auto-linked from the log) |
| Tax Deduction | `=TotalMiles*MileageRate` (IRS standard) |
| Savings | `=MIN(SavingsSaved/SavingsGoal,1)` |
| Driver Health Score | `=AVERAGE(HealthRange)` |

Every shift you log rolls into **live earnings, miles & fuel**; the Expenses
sheet pulls fuel straight from the log so your **cost-per-mile and true net**
stay honest; the Tax Center auto-picks **mileage vs actual** and banks a
quarterly set-aside; the Monthly Budget flows your **net income** into a plan;
and a **Driver Health Score** blends net earnings, $/hour, trips, savings,
consistency and your tax reserve.

**Verified sample driver** (Jordan, a full-time Austin driver, 16 shifts/mo):
Gross **$4,226** · Net **$2,786** · Net **$21.68/hr** · Net **$0.97/mi** ·
**2,871** miles · **316** trips · **128.5** hours · Expenses **$1,440** ·
Mileage deduction **$2,010** · Savings **70%** · Driver Health **90%**.

---

## Premium driver-software design

- Two-row **gold-divider headers** on every tab
- Gold-topped white **KPI cards** + a true driver dashboard (12 KPIs +
  net-earnings-trend, earnings-mix, best-shifts & expense charts)
- Platform-coded shift log with earnings data-bars; surge & bonus statuses
  color-code; a Tax Center deduction checklist so nothing gets left on the table
- Image-placeholder **Receipts & Vehicle Gallery** (Insert ▸ Picture-in-cell or
  `=IMAGE()`) for painless tax-time records
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Lyft_Driver_Command_Center.xlsx
python3 build_marketing.py                 # -> ../marketing/01..06.png
python3 build_marketing_detail.py          # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

# Travel Command Center™ — The Ultimate Travel Budget & Trip Planning System

> Not a travel spreadsheet — a **complete travel operating system**. One
> premium Excel & Google Sheets dashboard for budget, itinerary, bookings,
> packing, documents & memories — every trip, automatically organized.

| | |
| - | - |
| **Product** | Travel Command Center™ |
| **Target** | Couples & families · solo travelers · group-trip organizers · digital nomads · honeymoon & bucket-list planners |
| **Angle** | Know the true cost *before* you go, and land with every booking, document & packing list in one place. |
| **Formats** | Excel `.xlsx` (21-tab system + Welcome) + Google Sheets edition |
| **Pricing** | $24 single · **$34 bundle** · $49 with trip-planning guide · $99 agency/creator license |

---

## Contents

```
products/travel-command-center/
├── README.md
├── Travel_Command_Center.xlsx      ← Excel master (21-tab system + Welcome)
├── GOOGLE_SHEETS.md
├── BUILD_INSTRUCTIONS.md
├── ETSY_LISTING.md
└── build/
    ├── build_xlsx.py
    └── build_marketing.py
```

---

## The 21-sheet system

| # | Sheet | # | Sheet |
| - | ----- | - | ----- |
| 1 | Executive Travel Dashboard | 12 | Document Vault |
| 2 | Trip Profile | 13 | Currency Tracker |
| 3 | Master Itinerary | 14 | Savings Planner |
| 4 | Budget Command Center | 15 | Road Trip Planner |
| 5 | Expense Tracker | 16 | Group Travel (split the bill) |
| 6 | Flight Manager | 17 | Checklists |
| 7 | Accommodation | 18 | Memory Gallery |
| 8 | Transportation | 19 | Travel Journal |
| 9 | Activity Planner | 20 | Analytics Dashboard |
| 10 | Restaurants | 21 | Settings |
| 11 | Packing Center | | |

*(+ a Welcome / Start-Here tab — 22 tabs total.)*

---

## Signature automation (12 dashboard KPIs)

| Metric | How it's computed |
| ------ | ----------------- |
| Days to Departure | `=MAX(DepartDate-TODAY(),0)` |
| Total Budget | `=BudgetTotalPlanned` |
| Budget Left | `=BudgetTotalPlanned-BudgetTotalActual` |
| Spent So Far | `=BudgetTotalActual` (expenses roll in via `SUMIF`) |
| Flights Booked | `=COUNTA(FlightAirline)` |
| Hotels Confirmed | `=COUNTA(HotelName)` |
| Activities Booked | `=COUNTIF(ActStatus,"Booked")` |
| Countries | `=COUNTA(TripCountries)` |
| Docs Ready | `=COUNTIF(DocReady,"Yes")/COUNTA(DocName)` |
| Packing % | `=COUNTIF(PackDone,"Yes")/COUNTA(PackDone)` |
| Savings Goal | `=SavProgress` |
| Trip Readiness | `=AVERAGE(Analytics!C7:C11)` |

Every expense you log **auto-updates the budget actuals** (via `SUMIF` per
category), which flows straight into the dashboard's spending donut and
budget-vs-actual chart. **35 named ranges**, blank-safe `IFERROR` formulas,
cleanly-placed charts, and a **Trip Readiness score** on the Analytics tab.

---

## Premium travel-software design

- Two-row **gold-divider headers** on every tab
- Gold-topped white **KPI cards** + a true executive dashboard (12 KPIs +
  spending, budget, readiness & savings charts)
- Budget rows **data-bar** by % used; over-budget & unbooked items flag gold
- Image-placeholder **Memory Gallery** for your favorite trip photos
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

---

## Build & ship

```bash
cd build && python3 build_xlsx.py        # -> ../Travel_Command_Center.xlsx
python3 build_marketing.py               # -> ../marketing/01..06.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

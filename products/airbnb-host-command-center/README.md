# Airbnb Host Command Center™ — The Ultimate Airbnb Business Management System

> Not a booking spreadsheet — a **vacation-rental operating system**. One
> premium Excel & Google Sheets dashboard for reservations, finances, pricing,
> guests, cleaning, maintenance, inventory, reviews, taxes & multi-property
> analytics. Run your short-term rental like the business it is.

| | |
| - | - |
| **Product** | Airbnb Host Command Center™ |
| **Target** | New & experienced Airbnb hosts · vacation-rental owners · STR investors · co-hosts & property managers · cabin/beach-house/tiny-home hosts |
| **Angle** | Replace a tangle of disconnected files with one system — save time, improve organization, grow profit. |
| **Formats** | Excel `.xlsx` (19-tab system + Welcome) + Google Sheets edition |
| **Pricing** | $29 single · **$44 bundle** · $59 with host playbook · $129 agency/creator license |

---

## Contents

```
products/airbnb-host-command-center/
├── README.md
├── Airbnb_Host_Command_Center.xlsx   ← Excel master (19-tab system + Welcome)
├── GOOGLE_SHEETS.md
├── BUILD_INSTRUCTIONS.md
├── ETSY_LISTING.md
└── build/
    ├── build_xlsx.py
    └── build_marketing.py
```

---

## The 19-sheet system

| # | Sheet | # | Sheet |
| - | ----- | - | ----- |
| 1 | Executive Host Dashboard | 11 | Review Tracker |
| 2 | Property Profile | 12 | Message Template Library |
| 3 | Master Booking Calendar | 13 | Supplier Directory |
| 4 | Reservation Manager | 14 | Tax & Expense Preparation |
| 5 | Financial Command Center | 15 | Multi-Property Dashboard |
| 6 | Pricing Strategy | 16 | Goal Planner |
| 7 | Guest CRM | 17 | Photo & Improvement Planner |
| 8 | Cleaning Command Center | 18 | Analytics Dashboard |
| 9 | Maintenance Manager | 19 | Settings |
| 10 | Inventory Manager | | |

*(+ a Welcome / Start-Here tab — 20 tabs total.)*

---

## Signature automation (12 dashboard KPIs)

| Metric | How it's computed |
| ------ | ----------------- |
| Monthly Revenue | `=FinRevenue` (income roll-up) |
| Net Profit | `=FinNetProfit` (revenue − expenses) |
| Occupancy Rate | `=SUM(CalNights)/(properties×30)` |
| Avg Nightly Rate | `=FinNightly/SUM(CalNights)` (ADR) |
| Upcoming Check-Ins | `=COUNTIFS(CalCheckIn, next 7 days)` |
| Upcoming Check-Outs | `=COUNTIFS(CalCheckOut, next 7 days)` |
| Avg Length of Stay | `=SUM(CalNights)/COUNT(CalNights)` |
| Guest Rating | `=AVERAGE(RevOverall)` |
| Cleaning Pending | `=COUNTIF(CleanStatus,"Scheduled")` |
| Maintenance Due | `=COUNTIF(MaintStatus,"Open"/"In Progress")` |
| Low-Stock Items | `=SUMPRODUCT((InvQty<=InvReorder)*…)` |
| Business Health | `=AVERAGE(Analytics!C7:C12)` |

Every expense you log rolls into a live **P&L** (net profit, margin, cash
flow); the booking calendar drives occupancy, ADR & length of stay; inventory
fires **low-stock alerts**; and a **Business Health Score** blends occupancy,
margin, guest rating, maintenance, inventory & goals. **59 named ranges**,
blank-safe `IFERROR` formulas, cleanly-placed charts.

---

## Premium vacation-rental-software design

- Two-row **gold-divider headers** on every tab
- Gold-topped white **KPI cards** + a true executive dashboard (12 KPIs +
  expense, revenue-by-property, occupancy & booking-source charts)
- Budget & inventory rows **data-bar** by usage; REORDER & high-priority items
  flag red; scheduled turnovers flag gold; done/OK glow mint
- Image-placeholder **Photo & Improvement Planner** for listing upgrades
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

---

## Build & ship

```bash
cd build && python3 build_xlsx.py        # -> ../Airbnb_Host_Command_Center.xlsx
python3 build_marketing.py               # -> ../marketing/01..06.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

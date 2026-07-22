# Food Truck Command Center™ — The Complete Mobile-Food Business System

> Not a sales log — a **complete know-your-numbers mobile-food system**. One
> premium **Google Sheets + printable PDF** command center for running a truck:
> a menu & cost sheet, an **event P&L engine**, a break-even calculator, a daily
> sales log, commissary & overhead, inventory & par, a fuel & mileage log, a
> permits tracker, a supplies list, a bookings calendar and cash & tips.

| | |
| - | - |
| **Product** | Food Truck Command Center™ |
| **Target** | Food truck & trailer owners · concession & fair vendors · pop-ups & street food · coffee & dessert carts · new trucks finding their numbers · anyone selling food on wheels |
| **Angle** | Know your profit per event, cover your overhead & book the right gigs. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $19 single · **$29 bundle** (Sheets + PDF) · $39 with the recipe-costing add-on · $99 multi-truck / commercial license |

---

## Contents

```
products/foodtruck-command-center/
├── README.md
├── Food_Truck_Command_Center.xlsx    ← Google Sheets / Excel master (14 tabs)
├── Food_Truck_Printables.pdf         ← 12-page print-ready pack (US Letter)
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
| 2 | Dashboard | 9 | Fuel & Mileage |
| 3 | Menu & Cost | 10 | Permits |
| 4 | Events (P&L) | 11 | Supplies |
| 5 | Break-Even | 12 | Bookings |
| 6 | Daily Sales | 13 | Cash & Tips |
| 7 | Commissary & Overhead | 14 | Settings |

## The 12 printable PDF pages

Event P&L Sheet · Break-Even Worksheet · Daily Sales Log · Menu & Cost Card ·
Inventory & Par · Prep List · Fuel & Mileage Log · Permit Tracker · Supplies /
Shopping List · Bookings Calendar · Cash & Tips Reconciliation · Monthly P&L
Summary.

---

## Signature automation — the profit of every gig, and your break-even

Each event nets out **sales − food − fuel − fees − staff = net profit**, and the
break-even calculator divides your monthly fixed overhead by your average net per
event to tell you exactly how many gigs cover your costs:

```
Net profit / event        = Sales − Food − Fuel − Fee − Staff
Break-even events / month  = OverheadTotal ÷ Average(net profit per event)
```

### The 12 dashboard KPIs
Events · Total Sales · Total Profit · Avg Profit/Event · Food Cost · Avg
Sales/Event · Break-Even · Overhead/Mo · Top Event · Best Sales · Permits OK ·
Truck Score. The **Truck Score** blends profit margin, food-cost-on-target,
events-booked, permits-current, profitable-events and overhead-covered into one
0–100% number.

**Verified sample truck** (Rolling Smoke BBQ, owner Dana): **8** events · total
sales **$15,230** · total profit **$7,566** · avg profit **$946**/event · food
cost **29%** · avg sales **$1,904**/event · break-even **2.3 events** · overhead
**$2,140**/mo · top event **Corporate Catering** ($1,695 net) · best sales
**Music Festival** ($3,400) · permits **4/5** · **Truck Score 93%**.

---

## Premium mobile-food-software design

- An **event P&L engine** that nets out every gig and ranks them by profit
- A **break-even** calculator tied to your real commissary & overhead costs
- **Menu costing, daily sales, inventory & par**, and a **fuel & mileage** log
- **Permits**, **bookings** and **cash & tips** tools that keep the wheels turning
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A business tool, not financial, tax or legal advice.** Confirm figures and
> permit rules with the proper authorities.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Food_Truck_Command_Center.xlsx
python3 build_pdf.py                        # -> ../Food_Truck_Printables.pdf (+ page PNGs)
python3 build_marketing.py                 # -> ../marketing/01..06.png
python3 build_marketing_detail.py          # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

# Trucking Owner-Operator Command Center™ — Know Your Cost Per Mile

> Not a mileage log — a **complete cost-it, book-it, keep-it system**.
> One premium **Google Sheets + printable PDF** command center for owner-operators: a
> cost-per-mile engine (fixed ÷ miles + variable per mile, **divided by loaded miles**
> → your true cost and the rate you must never go below), fixed and variable costs, a
> load log with deadhead, settlements, a fuel log with live MPG, maintenance by odometer,
> truck & trailer, IFTA miles by state, a maintenance reserve fund and a monthly summary
> — everything cross-linked and live.

| | |
| - | - |
| **Product** | Trucking Owner-Operator Command Center™ |
| **Target** | Owner-operators (leased or under own authority) · new CDL holders buying a first truck · hot shot & box truck operators · small fleets running 2–5 trucks · drivers deciding whether to go solo |
| **Angle** | Deadhead isn't free — and your rate floor is lower than you think it is. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $27 single file · **$39 bundle** (Sheets + PDF) · $65 with a commercial-use file license — all the *same digital files*, never a service |

---

## Contents

```
products/trucking-command-center/
├── README.md
├── Trucking_Command_Center.xlsx     ← Google Sheets / Excel master (14 tabs)
├── Trucking_Printables.pdf          ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 8 | Fuel Log |
| 2 | Dashboard | 9 | Maintenance |
| 3 | Cost Per Mile | 10 | Truck & Trailer |
| 4 | Fixed Costs | 11 | IFTA & Miles |
| 5 | Variable Costs | 12 | Reserve Fund |
| 6 | Loads | 13 | Monthly Summary |
| 7 | Settlements | 14 | Settings |

## The 12 printable PDF pages

Cost Per Mile Worksheet · Fixed Costs Worksheet · Should I Take This Load? · Trip Sheet ·
Fuel Log · Maintenance & PM Schedule · Settlement Reconciliation · IFTA Miles by State ·
Pre-Trip Inspection · Maintenance Reserve Fund · Monthly Summary · Rate Floor Card.

---

## Signature automation — the cost-per-mile engine

There is one number that decides whether an owner-operator makes money, and most
calculators get it wrong by ignoring the empty miles:

```
Fixed cost per mile = fixed monthly costs ÷ TOTAL miles run
Variable per mile   = (diesel ÷ MPG) + maintenance + tires + tolls
Cost per mile run   = fixed per mile + variable per mile
Total cost          = cost per mile run × TOTAL miles
COST PER LOADED MILE = total cost ÷ LOADED miles     ← your rate floor
Profit per mile     = rate per loaded mile − cost per loaded mile
```

Your **costs run on total miles**, but only the **loaded** miles ever paid you. On the
sample month that gap is the whole story: **11,200 miles driven, 10,000 of them paid**.
A **$2.35** load is really a **$2.10** load once the 1,200 deadhead miles are charged to
it — and every load in the Loads tab is scored against the $1.41 floor automatically, with
anything below it flagged red.

### The 12 dashboard KPIs
Fixed/Month · Fixed Cost/Mile · Variable/Mile · Cost Per Mile Run · Cost/Loaded Mile ·
Rate/Loaded Mile · Profit/Loaded Mile · Loaded Miles · Deadhead % · Monthly Revenue ·
Monthly Profit · Road Score.
The **Road Score** blends profit per mile, net margin, deadhead, rate-vs-cost cover, fuel
economy and the maintenance reserve into one 0–100% number.

**Verified sample carrier** (Redline Freight Co., owner-operator Wes): fixed **$4,490**/mo
· 10,000 loaded + 1,200 deadhead = **11,200** miles · fixed **$0.401**/mi · variable
**$0.860**/mi · **cost per mile run $1.261** · total cost **$14,122** · **cost per loaded
mile $1.41** · rate **$2.35** · **profit $0.938/mi** · revenue **$23,500** · profit
**$9,378** (**39.9%** margin) · deadhead **10.7%** · actual MPG **6.5** · **Road Score
90%** (the honest weak spot: a maintenance reserve only 40% funded).

---

## Premium owner-operator design

- A **rate floor** you can defend on the phone with a broker
- **Deadhead charged**, not ignored — the single most-skipped cost in the industry
- **Every load scored** against your floor, losers flagged red automatically
- A **fuel log that computes actual MPG** from odometer readings, not a guess
- **IFTA miles by state**, **settlement deductions** and a **maintenance reserve fund**
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A business & organizing tool, not financial, tax, legal or DOT-compliance advice.**
> The pre-trip inspection page is an organizing aid, not a substitute for your carrier's
> required DVIR.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Trucking_Command_Center.xlsx
python3 build_pdf.py                        # -> ../Trucking_Printables.pdf (+ page PNGs)
python3 build_marketing.py                  # -> ../marketing/01..06.png
python3 build_marketing_detail.py           # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

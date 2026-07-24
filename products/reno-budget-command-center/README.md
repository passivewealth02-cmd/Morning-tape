# Home Renovation & Remodel Budget Command Center™ — The Remodel Operating System

> Not a budget sheet — a **complete plan-it, track-it, finish-it system**.
> One premium **Google Sheets + printable PDF** command center for a renovation: a
> budget-vs-actual engine (by room, with a contingency reserve), rooms, line items,
> contractors, payments, change orders, materials, a timeline, financing, decisions and
> a monthly summary — everything cross-linked and live.

| | |
| - | - |
| **Product** | Home Renovation & Remodel Budget Command Center™ |
| **Target** | Homeowners renovating · kitchen & bath remodels · whole-home renovations · DIY & GC-managed projects · house flippers & investors |
| **Angle** | Know what it costs, what you've spent, and what's left — before it's a surprise. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $19 single file · **$29 bundle** (Sheets + PDF) · $49 with a commercial-use file license — all the *same digital files*, never a service |

---

## Contents

```
products/reno-budget-command-center/
├── README.md
├── Reno_Budget_Command_Center.xlsx    ← Google Sheets / Excel master (14 tabs)
├── Reno_Budget_Printables.pdf         ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 8 | Change Orders |
| 2 | Dashboard | 9 | Materials |
| 3 | Budget vs Actual | 10 | Timeline |
| 4 | Rooms | 11 | Financing |
| 5 | Line Items | 12 | Decisions |
| 6 | Contractors | 13 | Monthly Summary |
| 7 | Payments | 14 | Settings |

## The 12 printable PDF pages

Budget Worksheet · Room Budget · Line Items · Contractor List · Payment Log · Change
Orders · Materials · Timeline · Financing Plan · Decisions Log · Monthly Summary ·
Remodel Checklist.

---

## Signature automation — budget vs actual, by room

Everything connects. Your room budgets set the total, your spending sets the actual, and
a contingency reserve absorbs the surprises:

```
Total budget    = Σ room budgets
Remaining       = total budget − total spent
Contingency     = contingency rate × total budget   (reserve − used = left)
Outstanding     = committed spend − paid to date
```

### The 12 dashboard KPIs
Total Budget · Total Spent · Remaining · Budget Used · Contingency · Contingency Left ·
Rooms · Paid To Date · Outstanding · Change Orders · % Complete · Reno Score.
The **Reno Score** blends under-budget, rooms-on-budget, contingency-healthy,
scope-defined, payments-current and change-orders-in-check into one 0–100% number.

**Verified sample remodel** (Cedar & Stone, homeowner Reese): total budget **$60,000**
· spent **$42,000** · remaining **$18,000** · budget used **70%** · contingency
**$9,000** · contingency left **$7,000** · rooms **5** · paid to date **$38,000** ·
outstanding **$4,000** · change orders **$3,000** · % complete **65%** · **Reno Score
90%**.

---

## Premium remodel design

- A **budget-vs-actual** engine with a **contingency reserve**
- **Budget by room** with spent, remaining and % used on each
- **Line items**, **contractors**, **payments** and **change orders**
- **Materials & finishes**, a **timeline**, **financing** and **decisions**
- A monthly **spend trend** and a live **Reno Score**
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A budgeting & organizing tool, not construction, financial or legal advice.**
> Confirm figures with your own contractors and advisors.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Reno_Budget_Command_Center.xlsx
python3 build_pdf.py                         # -> ../Reno_Budget_Printables.pdf (+ page PNGs)
python3 build_marketing.py                  # -> ../marketing/01..06.png
python3 build_marketing_detail.py           # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

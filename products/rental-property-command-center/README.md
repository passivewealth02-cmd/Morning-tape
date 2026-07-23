# Rental Property & Landlord Command Center™ — The Buy-&-Hold Operating System

> Not a rent tracker — a **complete analyze-it, run-it, measure-it system**.
> One premium **Google Sheets + printable PDF** command center for buy-and-hold
> landlords: a deal analyzer (rent − operating costs − mortgage → NOI, cash flow, cap
> rate, cash-on-cash, DSCR), a rent roll, tenants, a rent ledger, expenses,
> maintenance & CapEx, mortgages, reserves, mileage, renewals and a monthly summary —
> everything cross-linked and live.

| | |
| - | - |
| **Product** | Rental Property & Landlord Command Center™ |
| **Target** | Buy-and-hold landlords · first-time & aspiring investors · house-hackers · small multifamily owners · out-of-state investors · anyone analyzing a rental |
| **Angle** | Know the real return on every door. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $19 single file · **$29 bundle** (Sheets + PDF) · $49 with a commercial-use file license — all the *same digital files*, never a service |

---

## Contents

```
products/rental-property-command-center/
├── README.md
├── Rental_Property_Command_Center.xlsx  ← Google Sheets / Excel master (14 tabs)
├── Rental_Property_Printables.pdf       ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 8 | Maintenance & CapEx |
| 2 | Dashboard | 9 | Mortgage & Loan |
| 3 | Deal Analyzer | 10 | Reserves & Escrow |
| 4 | Rent Roll | 11 | Mileage & Travel |
| 5 | Tenants | 12 | Renewals & Docs |
| 6 | Rent Ledger | 13 | Monthly Summary |
| 7 | Expenses Log | 14 | Settings |

## The 12 printable PDF pages

Deal Analyzer · Rent Roll · Tenant Sheet · Rent Ledger · Expense Log · Maintenance &
CapEx · Mortgage & Loans · Reserves Tracker · Mileage Log · Renewals & Docs · Monthly
Summary · Move-In / Move-Out.

---

## Signature automation — the real return on a door

Everything connects. Your rent and expenses feed the NOI, the mortgage nets the cash
flow, and the purchase numbers turn it into the returns that decide buy-or-pass:

```
Gross income  = Rent + other income
NOI           = Gross income − operating expenses
Cash flow     = NOI − mortgage (P&I)
Cap rate      = NOI × 12 ÷ purchase price
Cash-on-cash  = Cash flow × 12 ÷ cash invested
DSCR          = NOI ÷ mortgage
```

### The 12 dashboard KPIs
Monthly Rent · Gross Income · Operating Exp · NOI · Mortgage · Cash Flow · Cap Rate ·
Cash-on-Cash · DSCR · 1% Rule · Annual Cash Flow · Landlord Score. The **Landlord
Score** blends cash-flow-positive, cap-rate, cash-on-cash, DSCR, expenses-in-check
and reserves-funded into one 0–100% number.

**Verified sample deal** (Maple Lane Rentals, 123 Maple St · A): rent **$1,850** ·
gross **$1,925** · operating exp **$830** · NOI **$1,095** · mortgage **$735** · cash
flow **$360** · cap rate **6.0%** · cash-on-cash **8.0%** · DSCR **1.49** · 1% rule
**0.84%** · annual cash flow **$4,320** · **Landlord Score 90%**.

---

## Premium landlord design

- A **deal analyzer** with NOI, cash flow, cap rate, cash-on-cash & DSCR
- A **rent roll** across every door + tenant, lease & deposit records
- A **rent ledger** and category **expense** and **CapEx** logs for taxes
- **Mortgage**, **reserves** and a **mileage** log (deductible)
- **Renewals & inspections** reminders and a **monthly summary**
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A business tool, not financial, legal, tax or investment advice.** Confirm every
> figure with your own advisors before you buy or refinance.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Rental_Property_Command_Center.xlsx
python3 build_pdf.py                         # -> ../Rental_Property_Printables.pdf (+ page PNGs)
python3 build_marketing.py                  # -> ../marketing/01..06.png
python3 build_marketing_detail.py           # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

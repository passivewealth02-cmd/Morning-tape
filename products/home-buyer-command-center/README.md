# First-Time Home Buyer & Mortgage Command Center™ — The Home-Buying Operating System

> Not a mortgage calculator — a **complete afford-it, find-it, close-it system**.
> One premium **Google Sheets + printable PDF** command center for buying your first
> home: an affordability engine (price + down payment + rate → your true all-in monthly
> payment and the DTI lenders check), down-payment savings, closing costs, a home
> comparison scorer, a lender compare, amortization, a life-after-buying budget, credit
> prep, a house hunting log, a moving checklist and a monthly summary — everything
> cross-linked and live.

| | |
| - | - |
| **Product** | First-Time Home Buyer & Mortgage Command Center™ |
| **Target** | First-time home buyers · couples buying together · renters ready to own · anyone saving a down payment · buyers comparing lenders |
| **Angle** | Know what you can truly afford — before an agent tells you. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $19 single file · **$29 bundle** (Sheets + PDF) · $49 with a commercial-use file license — all the *same digital files*, never a service |

---

## Contents

```
products/home-buyer-command-center/
├── README.md
├── Home_Buyer_Command_Center.xlsx    ← Google Sheets / Excel master (14 tabs)
├── Home_Buyer_Printables.pdf         ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 8 | Amortization |
| 2 | Dashboard | 9 | Life After Buying |
| 3 | Affordability | 10 | Credit Prep |
| 4 | Down Payment | 11 | House Hunting |
| 5 | Closing Costs | 12 | Moving Checklist |
| 6 | Home Comparison | 13 | Monthly Summary |
| 7 | Lender Compare | 14 | Settings |

## The 12 printable PDF pages

Affordability Worksheet · Home Comparison · Closing Cost Checklist · Down Payment
Tracker · Lender Comparison · House Hunting Log · Must-Have List · Credit Prep · Life
After Buying · Moving Checklist · Amortization Snapshot · Closing Day Checklist.

---

## Signature automation — your true monthly payment

Everything connects. A price and a down payment set the loan, the loan sets the
payment, and the payment gets checked against the ratios lenders actually use:

```
Loan amount   = price − (price × down %)
P&I           = PMT(rate/12, years×12, loan)
Monthly PITI  = P&I + property tax + insurance + PMI + HOA
Front-end DTI = PITI ÷ gross monthly income          (target ≤ 28%)
Back-end DTI  = (PITI + other debts) ÷ gross monthly (target ≤ 36%)
Cash to close = down payment + closing costs
```

### The 12 dashboard KPIs
Home Price · Down Payment · Loan Amount · Principal & Interest · Monthly Payment ·
Front-End DTI · Back-End DTI · Closing Costs · Cash to Close · Saved · Total Interest ·
Buyer Score. The **Buyer Score** blends cash-saved, payment-affordable, debt-healthy,
credit-ready, emergency-fund and 20%-down into one 0–100% number.

**Verified sample buyer** (Keystone & Co., Avery): home price **$290,000** · down
payment **$23,200** (8%) · loan **$266,800** · P&I **$1,686** · monthly payment
**$2,213** · front-end DTI **28.0%** · back-end DTI **33.6%** · closing costs **$8,700**
· cash to close **$31,900** · saved **$32,000** · total interest **$340,289** · **Buyer
Score 90%** (the honest weak spot: 8% down means PMI).

---

## Premium home-buying design

- An **affordability engine** with the full PITI — including PMI
- **Front- and back-end DTI** with a plain-English comfortable/stretched verdict
- **Down payment** savings with a months-to-funded projection
- **Home comparison** scoring and **lender** rate comparison
- **Credit prep**, a **life-after-buying** budget and a **moving** checklist
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A planning & organizing tool, not financial, lending, legal or real-estate advice.**
> Confirm every figure with your own lender and advisors.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Home_Buyer_Command_Center.xlsx
python3 build_pdf.py                         # -> ../Home_Buyer_Printables.pdf (+ page PNGs)
python3 build_marketing.py                  # -> ../marketing/01..06.png
python3 build_marketing_detail.py           # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

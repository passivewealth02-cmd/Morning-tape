# Home Renovation & Flip Command Center™ — The Ultimate House-Flipping System

> Not a rehab checklist — a **complete analyze-it, fund-it, flip-it operating
> system**. One premium **Google Sheets + printable PDF** command center with a
> deal analyzer (70% rule, all-in cost, cash-on-cash ROI & projected profit), a
> rehab budget (planned vs actual by category), scope of work, contractor
> directory, draw schedule, materials list, timeline, holding costs, financing,
> comps & ARV, a selling/exit net sheet, punch list and a before/after photo log.

| | |
| - | - |
| **Product** | Home Renovation & Flip Command Center™ |
| **Target** | New & experienced house flippers · BRRRR & buy-hold investors · wholesalers analyzing deals · DIY & whole-home renovators · investor-friendly agents · anyone renovating to sell or refinance |
| **Angle** | Know your number before you buy — then protect it to the closing table. |
| **Formats** | Google Sheets (17-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $19 single · **$29 bundle** (Sheets + PDF) · $39 investor / multi-deal · $99 team / commercial license |

---

## Contents

```
products/flip-command-center/
├── README.md
├── Flip_Command_Center.xlsx    ← Google Sheets / Excel master (17 tabs)
├── Flip_Printables.pdf         ← 12-page print-ready job-site binder (US Letter)
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

## The 17-tab system

| # | Tab | # | Tab |
| - | --- | - | --- |
| 1 | Start Here | 10 | Timeline |
| 2 | Dashboard | 11 | Holding Costs |
| 3 | Deal Analyzer | 12 | Financing |
| 4 | Property Details | 13 | Comps & ARV |
| 5 | Rehab Budget | 14 | Selling & Exit |
| 6 | Scope of Work | 15 | Punch List |
| 7 | Contractors | 16 | Photo Log |
| 8 | Draws & Payments | 17 | Settings |
| 9 | Materials | | |

## The 12 printable PDF pages

Deal Analyzer Worksheet · Rehab Budget · Scope of Work · Contractor & Bid Sheet ·
Draw Schedule · Materials Shopping List · Project Timeline · Holding-Cost
Worksheet · Comps & ARV Worksheet · Selling & Exit / Net Sheet · Punch List ·
Before & After Photo Log.

---

## Signature automation — the Deal Analyzer

Enter the deal once and everything computes:

| Output | Formula |
| ------ | ------- |
| Loan Amount | `=PurchasePrice*LoanLTV` |
| Down Payment | `=PurchasePrice-LoanAmount` |
| Selling Costs | `=ARV*SellCostPct` |
| All-In Cost | `=PurchasePrice+RehabBudget+BuyClosing+HoldingTotal+SellingCosts` |
| Cash Invested | `=DownPayment+RehabBudget+BuyClosing+HoldingTotal` |
| Projected Profit | `=ARV-AllInCost` |
| Cash-on-Cash ROI | `=ProjectedProfit/CashInvested` |
| 70% Rule — Max Offer | `=Rule70*ARV-RehabBudget` |
| Verdict | `=IF(PurchasePrice<=MAO70,"BUY","PASS")` |

### The 12 dashboard KPIs
ARV · Purchase · Rehab Budget · All-In Cost · Projected Profit · Cash-on-Cash
ROI · 70% Rule MAO · Verdict · Budget Used · Spent to Date · Tasks Done · Deal
Score. The **Deal Score** blends profit margin, ROI, the 70% rule, rehab
budget-vs-actual, scope completion and timeline into one 0–100% number, and a
**Rehab: Planned vs Actual** bar chart shows overruns at a glance.

**Verified sample deal** (the Maplewood Flip): ARV **$340,000** · Purchase
**$185,000** · Rehab **$45,000** · All-In **$266,802** · Projected Profit
**$73,198** · Cash-on-Cash ROI **77%** · 70% MAO **$193,000 → BUY** · Budget used
**76%** ($34K spent) · Scope **56%** done · **Deal Score 83%**.

---

## Premium investor-software design

- A real **deal analyzer** with the 70% rule and a straight BUY / PASS verdict
- **Rehab budget** with planned vs actual, % used, and red flags on any overrun
- The whole job coordinated: scope, contractors, draws, materials & a phased timeline
- Holding-cost and financing calculators that show what the money really costs
- Comps → ARV worksheet and a selling/exit **net sheet** for the closing table
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **Not financial, tax or investment advice.** All numbers are a sample deal —
> run your own, and confirm local costs, permits, financing and market comps.
> Real estate carries risk; past results don't guarantee future ones.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Flip_Command_Center.xlsx
python3 build_pdf.py                        # -> ../Flip_Printables.pdf (+ page PNGs)
python3 build_marketing.py                 # -> ../marketing/01..06.png
python3 build_marketing_detail.py          # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

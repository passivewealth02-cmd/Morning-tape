# Debt Payoff Command Center™ — The Complete Debt-Freedom System

> Not a debt list — a **complete pick-a-method, attack-and-win payoff system**.
> One premium **Google Sheets + printable PDF** command center for getting to
> $0: a debt-list engine, a snowball/avalanche payoff plan, a side-by-side method
> comparison, a payment log, a shrinking balance-history chart, a per-debt payoff
> order, an extra-payment finder, milestones, an interest tracker and
> accelerators — all rolling up into one live Payoff Momentum score.

| | |
| - | - |
| **Product** | Debt Payoff Command Center™ |
| **Target** | Anyone tackling credit-card debt · snowball & avalanche fans · couples paying off debt together · student-loan & car-loan payers · debt-free-journey & FIRE folks · anyone who wants a real plan |
| **Angle** | Pick a method, throw every extra dollar at it & watch your debt-free date arrive sooner. |
| **Formats** | Google Sheets (12-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $16 single · **$24 bundle** (Sheets + PDF) · $32 with the budget add-on · $79 coach / commercial license |

---

## Contents

```
products/debt-command-center/
├── README.md
├── Debt_Command_Center.xlsx      ← Google Sheets / Excel master (12 tabs)
├── Debt_Printables.pdf           ← 12-page print-ready pack (US Letter)
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

## The 12-tab system

| # | Tab | # | Tab |
| - | --- | - | --- |
| 1 | Start Here | 7 | Balance History |
| 2 | Dashboard | 8 | Extra Payment |
| 3 | Debts | 9 | Milestones |
| 4 | Payoff Plan | 10 | Interest Tracker |
| 5 | Snowball vs Avalanche | 11 | Accelerators |
| 6 | Payment Log | 12 | Settings |

## The 12 printable PDF pages

Debt List & Snapshot · Payoff Plan · Snowball vs Avalanche · Debt Payoff Tracker ·
Debt-Free Date Worksheet · Debt Thermometer · Interest Saved Worksheet ·
Extra-Payment Finder · Payment Log · Balance History · Milestones & Wins ·
Debt-Freedom Goals.

---

## Signature automation — a payoff plan that projects your debt-free date

A month-by-month payoff simulation (built into the workbook build) computes, for
**both** snowball (smallest balance first) and avalanche (highest rate first),
the months to debt-free, total interest and the exact debt-free date — plus the
interest you save by choosing avalanche. Minimums roll down a fixed attack order
as each debt is cleared.

### The 12 dashboard KPIs
Total Debt · Paid Off · % Paid · Monthly Payment · Extra Payment · Highest APR ·
Debt-Free Date · Months to Free · Total Interest · Interest Saved · Focus Debt ·
Momentum. The **Payoff Momentum Score** blends debt reduced, extra-payment
funded, on-time payments, milestones hit, per-debt progress and starter-fund
readiness into one 0–100% number.

**Verified sample household** (The Bennett Household, snowball, $300/mo extra):
Total debt **$47,800** (6 debts) · Paid off **$18,200** (**28%** of $66,000) ·
Monthly payment **$1,405** · Highest APR **27%** · Debt-free **Jan 2030** ·
Months to free **42** · Total interest **$9,908** · Interest saved by avalanche
**$873** · Paying $300 extra saves **$14,102** vs minimums-only · Focus debt
**Medical Bill** · **Momentum 62%**.

---

## Premium debt-payoff-software design

- A **debt-list engine** with per-debt % paid, a rate color scale & balance bars
- A **snowball/avalanche** comparison — months, interest & debt-free date for both
- A **payment log** with an on-time streak and a shrinking **balance-history** line
- An **extra-payment finder**, **milestones** and **accelerators** to shave off months
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A personal debt-payoff tool, not financial, tax or credit advice.** For big
> decisions, talk to a qualified professional.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Debt_Command_Center.xlsx
python3 build_pdf.py                        # -> ../Debt_Printables.pdf (+ page PNGs)
python3 build_marketing.py                 # -> ../marketing/01..06.png
python3 build_marketing_detail.py          # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

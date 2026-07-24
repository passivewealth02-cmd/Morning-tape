# Net Worth & FIRE Command Center™ — The Financial-Independence Operating System

> Not a net-worth sheet — a **complete know-your-number, get-there system**.
> One premium **Google Sheets + printable PDF** command center for the FIRE journey: a
> FIRE-number engine (annual spending → your number, progress, coast FIRE & years to
> FI), a net-worth roll-up, assets, liabilities, accounts, contributions, income &
> expenses, a savings-rate log, a coast & projection, milestones and a net-worth trend
> — everything cross-linked and live.

| | |
| - | - |
| **Product** | Net Worth & FIRE Command Center™ |
| **Target** | FIRE & early-retirement planners · new investors & savers · debt-payoff & net-worth trackers · couples planning together · anyone chasing financial freedom |
| **Angle** | Know your number, and your progress to it. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $19 single file · **$29 bundle** (Sheets + PDF) · $49 with a commercial-use file license — all the *same digital files*, never a service |

---

## Contents

```
products/networth-fire-command-center/
├── README.md
├── Networth_FIRE_Command_Center.xlsx   ← Google Sheets / Excel master (14 tabs)
├── Networth_FIRE_Printables.pdf        ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 8 | Contributions |
| 2 | Dashboard | 9 | Income & Expenses |
| 3 | Net Worth | 10 | Savings Rate |
| 4 | FIRE Number | 11 | Coast & Projection |
| 5 | Assets | 12 | Milestones |
| 6 | Liabilities | 13 | Net Worth Trend |
| 7 | Accounts | 14 | Settings |

## The 12 printable PDF pages

Net Worth Worksheet · FIRE Number Worksheet · Asset List · Liability List · Account
List · Contributions Log · Income & Expenses · Savings Rate Log · Coast & Projection ·
Milestones · Net Worth Trend · Money Checklist.

---

## Signature automation — know your number

Everything connects. Your spending sets your FIRE number, your assets and liabilities
set your net worth, and both give your progress and your years to freedom:

```
FIRE number    = annual spending ÷ safe withdrawal rate  (the 4% rule → ×25)
Net worth      = total assets − total liabilities
FIRE progress  = net worth ÷ FIRE number
Coast number   = FIRE number ÷ (1 + return)^(retire age − current age)
Years to FI    = NPER(return, −annual savings, −invested, FIRE number)
```

### The 12 dashboard KPIs
Net Worth · Total Assets · Total Liabilities · FIRE Number · FIRE Progress · Annual
Expenses · Savings Rate · Annual Savings · Coast Number · Years to FI · Invested Assets
· FIRE Score. The **FIRE Score** blends saving-enough, emergency-fund, positive-net-
worth, coast-FIRE, investing-to-goal and halfway-to-FI into one 0–100% number.

**Verified sample plan** (North Star Finance, Sage): net worth **$250,000** · assets
**$530,000** · liabilities **$280,000** · FIRE number **$1,000,000** · progress **25%**
· annual spending **$40,000** · savings rate **50%** · annual savings **$40,000** ·
coast number **$131,367** · years to FI **10.5** · invested **$200,000** · **FIRE Score
90%**.

---

## Premium FIRE design

- A **FIRE-number** engine (your annual spending → the amount you need)
- **Coast FIRE** and **years to FI** — when work becomes optional
- **Net worth** rolled up from assets & liabilities
- **Accounts**, **contributions** and a **savings-rate** log
- A **coast & projection** by age, **milestones** and a **net-worth trend**
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A personal-finance organizing tool, not financial, legal or tax advice.** Confirm
> figures with your own advisors.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Networth_FIRE_Command_Center.xlsx
python3 build_pdf.py                         # -> ../Networth_FIRE_Printables.pdf (+ page PNGs)
python3 build_marketing.py                  # -> ../marketing/01..06.png
python3 build_marketing_detail.py           # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

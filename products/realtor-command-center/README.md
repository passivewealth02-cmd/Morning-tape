# Real Estate Agent Command Center™ — The Realtor Operating System

> Not a commission calculator — a **complete net-it, track-it, hit-the-goal system**.
> One premium **Google Sheets + printable PDF** command center for real-estate agents:
> a commission calculator (sale price → GCI → split → net per deal), a deal pipeline,
> closings, buyers & sellers, listings, lead-source ROI, a database/sphere plan, a GCI
> goal tracker, business expenses, mileage and a monthly summary — everything
> cross-linked and live.

| | |
| - | - |
| **Product** | Real Estate Agent Command Center™ |
| **Target** | New & experienced agents · realtors & brokers · buyer's & listing agents · small teams & solo agents · referral & relocation agents · anyone tracking GCI & a goal |
| **Angle** | Know your GCI, your net, and your real business. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $19 single file · **$29 bundle** (Sheets + PDF) · $49 with a commercial-use file license — all the *same digital files*, never a service |

---

## Contents

```
products/realtor-command-center/
├── README.md
├── Realtor_Command_Center.xlsx        ← Google Sheets / Excel master (14 tabs)
├── Realtor_Printables.pdf             ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 8 | Lead Sources |
| 2 | Dashboard | 9 | Database |
| 3 | Commission Calculator | 10 | Goals & GCI |
| 4 | Pipeline | 11 | Business Expenses |
| 5 | Closings | 12 | Mileage & Auto |
| 6 | Buyers & Sellers | 13 | Monthly Summary |
| 7 | Listings | 14 | Settings |

## The 12 printable PDF pages

Commission Worksheet · Deal Pipeline · Closings Log · Buyers & Sellers · Listings ·
Lead Source ROI · Database Plan · GCI Goal Tracker · Business Expenses · Mileage Log ·
Monthly Summary · Closing Checklist.

---

## Signature automation — what a deal really nets

Everything connects. Your sale price and rate give the GCI, your split gives the agent
commission, and costs give the true net — then closings roll up to your GCI and net
income:

```
GCI (gross)     = Sale price × commission rate
Agent commission= GCI × agent split
Net per deal    = Agent commission − transaction costs
GCI (YTD)       = Σ agent commission on closings
Net income      = GCI (YTD) − business expenses
Goal progress   = GCI (YTD) ÷ annual GCI goal
```

### The 12 dashboard KPIs
Avg Sale Price · Commission · GCI/Deal · Net/Deal · Closings YTD · Volume YTD · GCI
YTD · Net Income · Pipeline · Pipeline GCI · Goal Progress · Agent Score. The **Agent
Score** blends net-margin, GCI-per-deal, closings-pace, pipeline, profit and
database-nurture into one 0–100% number.

**Verified sample agent** (Taylor Brooks, Summit Realty): avg sale **$450,000** ·
commission **3.0%** · GCI/deal **$13,500** · net/deal **$8,000** · **8** closings ·
volume **$3.6M** · GCI YTD **$75,600** · net income **$50,000** · **8** pipeline deals
· pipeline GCI **$30,000** · goal progress **76%** · **Agent Score 90%**.

---

## Premium realtor design

- A **commission calculator** (GCI, split, net per deal)
- A **deal pipeline** and a **closings** log → GCI & volume
- **Lead-source ROI**, a **database/sphere** plan and a **GCI goal** tracker
- **Buyers & sellers**, **listings**, **expenses** and a **mileage** log
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A business tool, not financial, legal or tax advice.** Confirm your split, fees
> and figures with your broker and your own advisors.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Realtor_Command_Center.xlsx
python3 build_pdf.py                         # -> ../Realtor_Printables.pdf (+ page PNGs)
python3 build_marketing.py                  # -> ../marketing/01..06.png
python3 build_marketing_detail.py           # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

# Contractor Job Costing & Bidding Command Center™ — Stop Bidding Blind

> Not an estimate template — a **complete bid-it, cost-it, close-it system**.
> One premium **Google Sheets + printable PDF** command center for contractors: a true
> bid engine (materials + labor + subs + equipment + overhead → the price that *actually*
> hits your margin), job costing against the bid, a job pipeline, crew hours, materials,
> subcontractors, change orders, equipment, invoices, a bid log and a monthly summary —
> everything cross-linked and live.

| | |
| - | - |
| **Product** | Contractor Job Costing & Bidding Command Center™ |
| **Target** | General contractors · remodelers · builders · handymen · painters, roofers, flooring & concrete crews · trade subs who bid their own work |
| **Angle** | Markup is not margin — and that gap is why your jobs come in short. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $27 single file · **$39 bundle** (Sheets + PDF) · $65 with a commercial-use file license — all the *same digital files*, never a service |

---

## Contents

```
products/contractor-command-center/
├── README.md
├── Contractor_Command_Center.xlsx     ← Google Sheets / Excel master (14 tabs)
├── Contractor_Printables.pdf          ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 8 | Subcontractors |
| 2 | Dashboard | 9 | Change Orders |
| 3 | Bid Builder | 10 | Equipment |
| 4 | Job Costing | 11 | Invoices |
| 5 | Jobs & Pipeline | 12 | Bid Log |
| 6 | Labor & Crew | 13 | Monthly Summary |
| 7 | Materials | 14 | Settings |

## The 12 printable PDF pages

Bid Worksheet · Job Cost Sheet · Materials Takeoff · Labor & Crew Log · Subcontractor Log
· Change Order Form · Daily Job Log · Bid Log · Invoice Schedule · Punch List · Monthly
Summary · Job Closeout.

---

## Signature automation — the bid engine (margin is a divisor, not a markup)

This is the whole product in four lines. Almost every contractor who loses money on a
"profitable" job made the same mistake: they *multiplied* by their margin instead of
*dividing* by what's left of the dollar.

```
Direct cost  = materials + labor hours × rate + subs + equipment
Overhead     = direct cost × overhead rate
Total cost   = direct cost + overhead
BID PRICE    = total cost ÷ (1 − margin target)      ← the correct formula
Planned profit = bid price − total cost
Actual margin  = (bid price − actual cost) ÷ bid price
```

The Bid Builder prints the wrong way right next to the right way, so the gap is
undeniable. On the sample job, marking cost up by 20% gives **$51,206** and a real margin
of only **16.7%** — **$2,134 less** than the correct $53,340 bid. Do that on 9 jobs a
year and it's roughly nineteen thousand dollars that never existed.

### The 12 dashboard KPIs
Direct Cost · Overhead · Total Cost · Bid Price · Planned Profit · Actual Cost · Actual
Profit · Actual Margin · Jobs · Backlog · Receivables · Builder Score.
The **Builder Score** blends job margin, overhead recovery, change-order capture,
receivables age, backlog depth and bid win rate into one 0–100% number.

**Verified sample company** (Ironwood Builders, owner Cal, sample job: full kitchen
remodel): materials **$14,000** · labor **320 hrs × $45 = $14,400** · subs **$8,500** ·
equipment **$1,200** → direct **$38,100** · overhead 12% **$4,572** · total cost
**$42,672** · **BID $53,340** · planned profit **$10,668** · actual cost $41,178 ·
actual profit **$12,162** · actual margin **22.8%** · 9 jobs · backlog **$186,000** ·
receivables **$31,000** · **Builder Score 90%** (the honest weak spot: a 40.9% bid win
rate — 9 wins from 22 bids).

---

## Premium contractor design

- A **bid engine** that divides by margin instead of marking up cost
- A printed **"markup is not margin"** comparison the buyer can't argue with
- **Job costing against the bid**, line by line, so overruns show while you can still fix them
- **Change orders tracked as revenue**, not favors — signed, priced and billed
- **Crew hours, burden, materials, subs and equipment** all feeding one number
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A business & organizing tool, not financial, tax, legal or accounting advice.** Labor
> burden, insurance, bonding and licensing costs vary by state and trade — enter your own
> in Settings.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Contractor_Command_Center.xlsx
python3 build_pdf.py                        # -> ../Contractor_Printables.pdf (+ page PNGs)
python3 build_marketing.py                  # -> ../marketing/01..06.png
python3 build_marketing_detail.py           # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

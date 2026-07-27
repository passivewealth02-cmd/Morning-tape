# Relationship & Couples Command Center™ — 50/50 Is Not the Same as Fair

> Not a shared budget — a **complete split-it, share-it, say-it system**.
> One premium **Google Sheets + printable PDF** command center for two people building a
> life: a **fair-share engine** (split the bills by what you each earn, not down the
> middle), shared bills, the **invisible labour split in hours**, money goals, savings,
> date nights, a weekly check-in, the big conversations, household admin, individual money
> and a monthly summary — everything cross-linked and live.

| | |
| - | - |
| **Product** | Relationship & Couples Command Center™ |
| **Target** | Couples moving in together · two people earning different amounts · newly engaged or newly married · anyone merging finances at all · anyone doing more than their share · couples who avoid money conversations |
| **Angle** | Same house. Same bills. Completely different life. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $22 single file · **$29 bundle** (Sheets + PDF) · $49 with a commercial-use file license — all the *same digital files*, never a service |

---

## Contents

```
products/couples-command-center/
├── README.md
├── Couples_Command_Center.xlsx     ← Google Sheets / Excel master (14 tabs)
├── Couples_Printables.pdf          ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 8 | Date Nights |
| 2 | Dashboard | 9 | Weekly Check-In |
| 3 | Fair Share | 10 | Big Conversations |
| 4 | Shared Bills | 11 | Household Admin |
| 5 | Invisible Labour | 12 | Individual Money |
| 6 | Money Goals | 13 | Monthly Summary |
| 7 | Savings | 14 | Settings |

## The 12 printable PDF pages

Fair Share Worksheet · Shared Bills · Invisible Labour · The Conversation · Money Goals ·
Weekly Check-In · The Big Conversations · Date Nights · Household Admin · Individual
Money · Month in Review · If Incomes Change.

---

## Signature automation — two engines

**1. The fair-share engine.** Two people earning different amounts splitting the bills
down the middle is not equality — it only looks like it:

```
Each person's share = their income ÷ combined income
Fair contribution   = shared bills × their share
What's left         = income − fair contribution
Kept %              = what's left ÷ their own income     ← identical for both
```

On the sample numbers ($4,200 and $6,300 against $3,908 of shared bills): **a 50/50 split
leaves one partner with $2,246 and the other with $4,346 — a $2,100 a month gap.** Split
proportionally and they each keep **exactly 62.8%** of their own income. Same house, same
bills, completely different life.

**2. The invisible labour split.** Hours a week, including the remembering:

```
Chore ratio      = more hours ÷ fewer hours
Extra hours/year = |difference| × 52
```

The sample comes out **23.0 vs 8.0 hours a week — a 2.88× split, and 780 extra hours a
year**, which is **19.5 full working weeks**. The workbook asks both partners to fill it
in *separately* and then compare, because the work you don't do is genuinely hard to see.

### The 12 dashboard KPIs
Shared Bills · 50/50 Leaves A · 50/50 Leaves B · The 50/50 Gap · A's Fair Share · B's Fair
Share · You Each Keep · How Close to Fair · Chore Split · Extra Hours a Year · Saved This
Month · Together Score.
The **Together Score** blends fair bills, check-ins, saving, goals, date nights and the
housework split into one 0–100% number — and it points at the **one** thing worth fixing.

**Verified sample couple** (Nadia & Sam): incomes **$4,200 / $6,300** (40% / 60%) · shared
bills **$3,908** · 50/50 leaves **$2,246 / $4,346** (**$2,100 gap**) · fair share
**$1,563.20 / $2,344.80** leaving both with **62.8%** · fairness **100%** · chores **23.0
vs 8.0 hrs** (**2.88×**, **780 extra hours a year**) · saved **$940** · **Together Score
90%** (the honest weak spot: housework at 40%).

---

## Design notes — this one could easily start a fight

It's built not to:

- The Start Here tab has a **"how to use this without a fight"** block
- Both partners fill in the labour tab **separately, then compare** — the instruction is
  explicit, and so is the line *"nobody is lying when the two columns don't match"*
- The printable has a whole page (**The Conversation**) that is only discussion prompts,
  with *"pick ONE row to move this month. Not all of them. One."*
- Dashboard status labels read "Good / OK / **Talk about it**" rather than "Watch"
- The **Individual Money** tab protects money that stays individual, and the
  no-questions-asked spending limit
- There is a plain safeguarding line: **if money is being controlled rather than shared, a
  spreadsheet is not the right help**
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`; the crest is **two interlocking rings**

> **An organizing tool, not relationship or financial advice.** It can't tell you what a
> fair life together looks like — only the two of you can.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Couples_Command_Center.xlsx
python3 build_pdf.py                        # -> ../Couples_Printables.pdf (+ page PNGs)
python3 build_marketing.py                  # -> ../marketing/01..06.png
python3 build_marketing_detail.py           # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

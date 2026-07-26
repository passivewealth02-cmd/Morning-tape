# Savings Challenge & Sinking Funds Command Center™ — The Save-On-Purpose Operating System

> Not a savings tracker — a **complete name-it, save-it, celebrate-it system**.
> One premium **Google Sheets + printable PDF** command center for saving on purpose: a
> sinking-funds engine (target ÷ 12 → what to set aside every month), the 100-Envelope
> Challenge, the 52-Week Challenge, custom challenges, goal countdowns, cash envelopes,
> savings accounts, an emergency fund, a deposit log, a savings streak and a monthly
> summary — everything cross-linked and live.

| | |
| - | - |
| **Product** | Savings Challenge & Sinking Funds Command Center™ |
| **Target** | Cash stuffers & envelope savers · anyone doing a savings challenge · budgeters tired of surprise bills · families saving for Christmas · anyone building an emergency fund |
| **Angle** | Give every saved dollar a name — and never be blindsided by a bill again. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $15 single file · **$22 bundle** (Sheets + PDF) · $39 with a commercial-use file license — all the *same digital files*, never a service |

---

## Contents

```
products/savings-challenge-command-center/
├── README.md
├── Savings_Challenge_Command_Center.xlsx   ← Google Sheets / Excel master (14 tabs)
├── Savings_Challenge_Printables.pdf         ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 8 | Savings Accounts |
| 2 | Dashboard | 9 | Emergency Fund |
| 3 | Sinking Funds | 10 | Deposit Log |
| 4 | 100 Envelope | 11 | Savings Streak |
| 5 | 52-Week | 12 | Monthly Summary |
| 6 | Challenges | 13 | Settings |
| 7 | Goal Countdown / Cash Envelopes | 14 | — |

*(Full order: Start Here · Dashboard · Sinking Funds · 100 Envelope · 52-Week ·
Challenges · Goal Countdown · Cash Envelopes · Savings Accounts · Emergency Fund ·
Deposit Log · Savings Streak · Monthly Summary · Settings)*

## The 12 printable PDF pages

Sinking Funds Worksheet · **100 Envelope Grid** · **52-Week Tracker** · Savings
Thermometer · Cash Envelope Tracker · Deposit Log · Goal Countdown · No-Spend Tracker ·
Savings Streak · Emergency Fund · Monthly Summary · Savings Checklist.

---

## Signature automation — the sinking-funds engine

Everything connects. Every yearly bill becomes a monthly number, and the calendar tells
you whether you're actually on pace:

```
Monthly set-aside = each fund's annual target ÷ 12   (summed across all funds)
Expected to date   = total target × months elapsed ÷ 12
On pace            = saved ÷ expected to date
100-Envelope total = 1 + 2 + … + 100 = $5,050
52-Week total      = 1 + 2 + … + 52  = $1,378
```

### The 12 dashboard KPIs
Funds Target · Funds Saved · Monthly Set-Aside · On Pace · Emergency Fund · Total Saved
· 100 Envelope · 52-Week · Active Challenges · Saved This Month · Savings Streak ·
Savings Score. The **Savings Score** blends funds-on-pace, emergency-fund,
challenges-running, monthly-goal, every-bill-covered and streak into one 0–100% number.

**Verified sample saver** (Copper & Clover, June): funds target **$9,000** · saved
**$6,000** · monthly set-aside **$750** · on pace **100%** · emergency fund **$6,000** ·
total saved **$12,274** · 100-Envelope **$5,050** · 52-Week **$1,378** · active
challenges **3** · saved this month **$520** · streak **24** days · **Savings Score 90%**
(the honest weak spot: a 24-day streak against a 60-day goal).

---

## Premium savings design

- A **sinking-funds engine** that turns yearly bills into one monthly number
- An **on-pace check** against the calendar, not just a running total
- The **100-Envelope** ($5,050) and **52-Week** ($1,378) challenges, built in
- **Cash-envelope stuffing**, a **savings streak** and a **no-spend** tracker
- A separate **emergency fund** — because Christmas is not an emergency
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A budgeting & organizing tool, not financial advice.**

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Savings_Challenge_Command_Center.xlsx
python3 build_pdf.py                         # -> ../Savings_Challenge_Printables.pdf (+ page PNGs)
python3 build_marketing.py                  # -> ../marketing/01..06.png
python3 build_marketing_detail.py           # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

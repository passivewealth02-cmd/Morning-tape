# Dating Life Command Center™ — See It Clearly

> Not a cute planner — a **complete count-it, score-it, decide system**.
> One premium **Google Sheets + printable PDF** command center for dating: a dating funnel
> (matches → conversations → first dates → second dates, with what each one really costs
> in time and money), an **effort & reciprocity scorecard**, green and red flag lists, the
> people you're seeing, a date log, conversations, time & money, non-negotiables, a safety
> plan, monthly reflection and a summary — everything cross-linked and live.

| | |
| - | - |
| **Product** | Dating Life Command Center™ |
| **Target** | Anyone dating on apps and tired of it · back out there after a long relationship · dating with intention this year · anyone who keeps picking the same person · anyone doing all the texting |
| **Angle** | Not to judge anyone. To see clearly what you already feel. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $18 single file · **$24 bundle** (Sheets + PDF) · $39 with a commercial-use file license — all the *same digital files*, never a service |

---

## Contents

```
products/dating-command-center/
├── README.md
├── Dating_Command_Center.xlsx     ← Google Sheets / Excel master (14 tabs)
├── Dating_Printables.pdf          ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 8 | Conversations |
| 2 | Dashboard | 9 | Time & Money |
| 3 | Dating Funnel | 10 | Non-Negotiables |
| 4 | Effort & Reciprocity | 11 | Safety Plan |
| 5 | Green & Red Flags | 12 | Reflection |
| 6 | People | 13 | Monthly Summary |
| 7 | Date Log | 14 | Settings |

## The 12 printable PDF pages

Non-Negotiables · The Month, Counted · Effort & Reciprocity · Green Flags · Red Flags ·
Date Log · People · Conversations · Time & Money · Safety Plan · Reflection · Month in
Review.

---

## Signature automation — two engines

**1. The dating funnel.** Dating is the only part of life we're told not to look at
clearly. Everywhere else you'd count:

```
Matches → conversations → first dates → second dates → still seeing
Hours    = swiping + messaging + (first dates × hours per date + prep)
Cost per second date  = total spend ÷ second dates
Hours per second date = total hours ÷ second dates
```

On the sample month: **240 matches became 5 second dates**, at **69 hours and $497** —
which means every second date cost about **$99.40 and 13.8 hours**. Not to make anyone
feel bad. To let them decide, on purpose, how they want to spend a Tuesday.

**2. The effort scorecard.** Six dimensions, scored out of ten for you and for them:

```
Effort ratio = your total ÷ their total      (1.0 is even)
```

The sample comes out at **48 vs 16 — you're doing 3× the work**. That single number is the
thing the buyer has felt for weeks and couldn't prove, and it is the reason this product
sells.

### The 12 dashboard KPIs
Matches · Conversations · First Dates · Second Dates · Second-Date Rate · Still Seeing ·
Hours · Spent · Per Second Date · Hours Per Second Date · Effort Ratio · Dating Score.
The **Dating Score** blends meeting-not-texting, second-date rate, flags, budget, safety
steps and reciprocity into one 0–100% number.

**Verified sample month**: 240 matches → 68 conversations (28.3%) → 14 first dates (5.8%
of matches) → 5 second dates (35.7%) → 2 still seeing · 9 days from match to first date ·
69 hours · $496.98 · **$99.40 and 13.8 hours per second date** · 9 green flags, 2 red ·
**effort ratio 3.0×** · **Dating Score 90%** (the honest weak spot: reciprocity at 40%).

---

## Design notes — tone matters more than usual here

This product could very easily be cruel. It isn't, deliberately:

- The score measures **your dating practice**, never your desirability
- The effort scorecard says out loud that **a low score doesn't make someone a bad
  person** — it means you're doing more of the work, and what you do about that is yours
- The printable effort page ends with three fairness prompts (*"Have I actually told them
  what I need, out loud, once?"*)
- The **Safety Plan** tab is not an afterthought — it's nine steps, a who-knows-where-you-
  are block, and the line *"You never owe anyone a second more of your evening."*
- The file label reads **"private · yours"**
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`; the crest is a **struck match**

> **A personal organizing and reflection tool** — not relationship, psychological or
> medical advice. No spreadsheet can tell anyone whether to love someone. The scores are a
> mirror, not a verdict.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Dating_Command_Center.xlsx
python3 build_pdf.py                        # -> ../Dating_Printables.pdf (+ page PNGs)
python3 build_marketing.py                  # -> ../marketing/01..06.png
python3 build_marketing_detail.py           # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

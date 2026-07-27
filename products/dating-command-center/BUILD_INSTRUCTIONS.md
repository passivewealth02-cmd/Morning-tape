# Dating Life Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/dating-command-center/build
python3 build_xlsx.py      # -> ../Dating_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** carries the tone-setting note and the disclaimer: a personal organizing
   and reflection tool, not relationship, psychological or medical advice; the scores are
   a mirror, not a verdict. **Do not cut this block** — it is what keeps the product kind.
2. **Dating Funnel**: 240 matches → 68 conversations → 14 first dates → 5 second dates →
   2 still seeing. **MATCH-TO-DATE 5.8%** (`MatchToDate`), **SECOND-DATE RATE 35.7%**
   (`SecondDateRate`).
3. Hours: 9 swiping + 11 messaging + 14 × (2.5 + 1.0) = **69 HOURS** (`HoursSpent`).
   Spend **$496.98** (`SpendTotal`). **PER FIRST DATE $35.50**, **PER SECOND DATE $99.40**
   (`CostPerSecondDate`), **13.8 HOURS PER SECOND DATE** (`HoursPerSecondDate`).
4. **Effort & Reciprocity**: six dimensions, you **48** (`YourEffort`) vs them **16**
   (`TheirEffort`) → **EFFORT RATIO 3.0×** (`EffortRatio`), and the verdict cell reads
   **CARRYING ALL OF IT**. Gaps of 5+ shade red. **This tab is the product's whole sales
   argument — check it renders.**
5. **Green & Red Flags**: two 12-item lists side by side; **9 green** (`GreenCount`),
   **2 red** (`RedCount`), **net 7** (`FlagNet`). Green "Yes" shades mint, red "Yes"
   shades red.
6. **Conversations** flags anyone whose share of the messages is under 35% — the sample
   has four such rows, all of them people she's never met.
7. **Safety Plan** counts **9** steps held (`SafetyHeld`); any "No" shades red.
8. **Time & Money** sums to **$496.98** across four lines and shows hours in a second
   block, including the "that's 2.9 full days" line.
9. **Dashboard** fills 12 KPI cards + a How-It's-Going table + a where-the-69-hours-went
   donut and a first-dates-by-month chart. **Dating Score 90%** (reciprocity is the honest
   weak dimension). Status labels read "Good / OK / **Look at this**" rather than
   "Strong / OK / Watch" — softer on purpose.

> Note: uses `SUM`, `COUNTA`, `COUNTIF`, `MIN`, `MAX`, `AVERAGE`, `IF`, `IFERROR` — opens
> in Google Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Dating_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter. Twelve
pages: **non-negotiables first** (deliberately page one — write them before you meet
anyone), the month counted, effort & reciprocity, a full green-flag page, a full red-flag
page, date log, people, conversations, time & money, safety plan, reflection prompts and a
month in review.

Page 3 (Effort & Reciprocity) ends with three fairness prompts so the page can't be used
purely as a weapon. Page 10 (Safety Plan) ends with *"You never owe anyone a second more
of your evening. Leaving early is always allowed."* Both of those lines stay.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with the where-the-69-hours-went donut and
first-dates chart), everything-inside (14 tabs), the **effort scorecard**, the dating
funnel, the clarity engine (both), and the **12-page printables showcase**. Images 3–5
each show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "a dating journal vs Command Center",
09 how it actually works in 4 steps, 10 what's-included / who-it's-for / works-with. Ten
images — fills all 10 Etsy slots. All headline numbers (240 matches · 68 conversations ·
14 first dates · 5 second dates · 35.7% · 9 days · 69 hours · $497 · $99.40 per second
date · 7 net flags · 3.0× effort · 90% score) are verified against the workbook, and the
hours donut splits to exactly 69 (35 / 14 / 11 / 9 hrs).

> This build adds an optional `fmt` callable to the shared `vbars` helper so bar labels
> can render plain counts instead of `$0k`. Carry that forward — several products in the
> women's list chart counts rather than money.

The crest is a **struck match** — the spark, before you decide whether it's worth keeping
lit. Every product in the women's/dating line gets its own crest; this one starts the set.

---

## D. Etsy delivery package

```
Dating_Command_Center.xlsx          ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
Dating_Printables.pdf               ← 12-page print-ready pack
START_HERE.pdf                      ← onboarding quick-start
```

---

## E. Pricing  (all SKUs are the same digital files — a product, never a service)

| SKU | What the buyer downloads | Price |
| --- | ------ | ----- |
| DAT-GS   | The Google Sheets / Excel file only | $18 |
| DAT-PDF  | The printable PDF only | $15 |
| DAT-BUNDLE | The spreadsheet + the printable PDF | **$24** |
| DAT-COMM | The same files + a commercial-use file license | $39 |

> ⚠ **Etsy Services-policy safety, and one extra rule for this niche.**
>
> No setup help, consultations, **coaching or mentoring**, "dating coaching", profile
> reviews, or "free updates / lifetime access". Plain digital file only. The dating niche
> is full of coaching offers and that is exactly the category Etsy removes.
>
> **Also: never position this as advice.** Not "how to find love", not "attract the right
> partner", not anything that reads as a psychological claim. It is an organizing and
> reflection tool. That framing is both honest and safe.

- **Huge, warm Etsy audience — this is the volume anchor of the women's line.** Dating
  journals and green-flag/red-flag printables sell in enormous numbers at $4–$10. Both
  flag lists are *included* here, so you beat that crowd on quantity and then offer two
  things none of them have: a funnel with real costs, and the effort ratio.
- **Demand peaks in January** (new year, back on the apps) and again in **September**
  (cuffing season). There's a smaller spike right after Valentine's Day.
- Use all 10 photos + a walkthrough video that shows the *file*. Lead photo = the
  feature-forward hero; **the effort scorecard is your single most persuasive image** —
  "you're doing 3× the work" is the sentence this buyer has been trying to say out loud.
- **Sell it as clarity, not strategy.** "See it clearly" and "you already know" outperform
  anything that sounds like a system for winning at dating. The whole product is built on
  that tone; don't undercut it in the ad copy.
- Cross-sell the rest of the women's line as it lands — Situationship Clarity, Self-Love
  & Healing, Relationship & Couples.

---

## F. Maintenance

- Edit the `MATCHES`, `CONVERSATIONS`, `FIRST_DATES`, `SECOND_DATES`, `STILL_SEEING`,
  `DAYS_TO_FIRST_DATE`, `SPEND_BUDGET`, `SWIPE_HOURS`, `MESSAGE_HOURS`, `HOURS_PER_DATE`,
  `PREP_HOURS_PER_DATE` constants and the `DAYS_GOAL`, `SECOND_DATE_GOAL`, `FLAG_GOAL`,
  `EFFORT_GOAL` targets plus the `SPENDING`, `FUNNEL_MONTHS`, `YOUR_EFFORT`,
  `THEIR_EFFORT`, `GREEN_FLAGS`, `RED_FLAGS`, `PEOPLE`, `DATE_LOG`, `CONVOS`,
  `NON_NEGOTIABLES`, `SAFETY`, `REFLECTION`, `MONTHS` tables in `build_xlsx.py`.
- **Keep `SPENDING` summing to the figure quoted in the marketing** ($496.98) and
  `FUNNEL_MONTHS`' last row matching the headline counts. They currently do.
- **If you ever soften the sample effort scores, re-check the Dating Score.** The 90% is
  built on reciprocity landing at exactly 0.40 (goal 1.2 ÷ ratio 3.0).
- Keep `build_marketing.py`'s KPIs in sync with the workbook.

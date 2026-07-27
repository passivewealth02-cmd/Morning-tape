# Relationship & Couples Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/couples-command-center/build
python3 build_xlsx.py      # -> ../Couples_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** carries the "how to use this without a fight" block and the
   safeguarding line about money being controlled rather than shared. **Do not cut
   either** — they are what make this product safe to sell.
2. **Shared Bills** sums to **$3,908** (`SharedBills`) across eight lines.
3. **Fair Share**: incomes $4,200 / $6,300 → shares **40% / 60%** (`ShareA`, `ShareB`).
   The 50/50 block shows **$1,954** each and leaves **$2,246 / $4,346** (`LeftHalfA`,
   `LeftHalfB`) — a **$2,100 gap** (`HalfGap`).
4. The proportional block shows **$1,563.20 / $2,344.80** (`FairA`, `FairB`) leaving both
   at **62.8%** of their own income (`KeepPctA`, `KeepPctB`). **Those two percentages must
   be identical — that is the entire point of the page.**
5. Fairness accuracy reads **100%** (`Fairness`) because the sample couple has adopted the
   proportional split.
6. **Invisible Labour**: 8 tasks, **23.0** (`HoursA`) vs **8.0** (`HoursB`) hours a week →
   **CHORE RATIO 2.88×** (`ChoreRatio`), **780 EXTRA HOURS A YEAR** (`ExtraHoursYear`) =
   **19.5 working weeks**. **This tab is the product's whole sales argument — check it
   renders.**
7. **Money Goals**: 5 goals, **5 on track or better** (`GoalsOnTrack` / `GoalsTotal`).
   **Savings** sums to **$940** a month (`SavedMonth`). **Date Nights** counts **4**
   (`DateNights`). **Weekly Check-In** counts **4** (`CheckinsDone`). **Big
   Conversations** shows **6 had, 2 still to have** (`TalksHad`).
8. **Dashboard** fills 12 KPI cards + a How-You're-Doing table + an hours donut and a
   saved-by-month chart. **Together Score 90%** (housework is the honest weak dimension).
   Status labels read "Good / OK / **Talk about it**".

> Note: uses `SUM`, `COUNTA`, `COUNTIF`, `MIN`, `MAX`, `ABS`, `AVERAGE`, `IF`, `IFERROR` —
> opens in Google Sheets or Excel 2019/365. Several header cells use `=PartnerA` /
> `=PartnerB` so the couple's real names propagate from Settings.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Couples_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter. Twelve
pages: fair-share worksheet, shared bills, invisible labour, **the conversation**, money
goals, weekly check-in, ten big conversations, date nights, household admin, individual
money, month in review and **if incomes change**.

**Page 4 (The Conversation) is the most important page in the pack.** It is deliberately
not a form — it's the four ground rules ("nobody is lying when the two columns don't
match", "pick ONE row to move this month") plus five discussion prompts. It's what stops
page 3 being used as a weapon. Keep it, and feature it in the printables showcase.

Page 12 exists because the fair split needs re-running whenever an income changes — that's
a five-minute conversation once, not a renegotiation.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with the hours donut and saved-by-month
chart), everything-inside (14 tabs), the **fair-share engine**, the invisible labour
split, the fairness engine (both), and the **12-page printables showcase**. Images 3–5
each show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "a shared budget vs Command Center",
09 how it actually works in 4 steps, 10 what's-included / who-it's-for / works-with. Ten
images — fills all 10 Etsy slots. All headline numbers ($3,908 bills · $2,246 / $4,346 ·
$2,100 gap · $1,563 / $2,345 fair · 62.8% each · 2.88× chore split · 780 extra hours ·
$940 saved · 90% score) are verified against the workbook.

The crest is **two interlocking rings** — gold and mint, genuinely linked, with bright
nodes where they meet. Note the geometry: rings need `off < rr` or they read as one shape.

> The dashboard KPI labels on this product are longer than usual ("50/50 LEAVES ONE OF
> YOU"), so the card label uses `fit_font` to shrink rather than overflow. Carry that
> forward wherever labels are long.

---

## D. Etsy delivery package

```
Couples_Command_Center.xlsx         ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
Couples_Printables.pdf              ← 12-page print-ready pack
START_HERE.pdf                      ← onboarding quick-start
```

---

## E. Pricing  (all SKUs are the same digital files — a product, never a service)

| SKU | What the buyer downloads | Price |
| --- | ------ | ----- |
| CPL-GS   | The Google Sheets / Excel file only | $22 |
| CPL-PDF  | The printable PDF only | $19 |
| CPL-BUNDLE | The spreadsheet + the printable PDF | **$29** |
| CPL-COMM | The same files + a commercial-use file license | $49 |

> ⚠ **Two rules on this listing.**
>
> **1. No services.** No setup help, consultations, **couples coaching or counselling**,
> or "free updates / lifetime access". Plain digital file only.
>
> **2. Never position it as therapy or advice.** Not "fix your relationship", not
> "communication method", nothing that reads as a clinical claim. It is an organizing
> tool. That framing is honest, safe, and converts better.

- **Two strong buyer moments, not one.** Couples moving in together, and couples merging
  finances after an engagement. Both are searched heavily and both are gift-able — this is
  one of the few products in the shop someone buys *for* a couple.
- **Demand peaks January** (new year, money resolutions) and **engagement season, December
  to February**. There's a real spike around Valentine's Day for the gift angle.
- Use all 10 photos + a walkthrough video that shows the *file*. Lead photo = the
  feature-forward hero; **the fair-share image is your single most persuasive one** —
  "50/50 quietly leaves a $2,100 a month gap" is a fact most couples have never had put in
  front of them.
- **The invisible labour angle is what makes this shareable.** "780 extra hours a year" is
  the number people screenshot. Lead one of your Etsy photos and any social post with it.
- Cross-sell **Dating Life** (the same buyer, one stage earlier) and the **Home Buyer**
  and **Wedding** products.

---

## F. Maintenance

- Edit the `A_NAME`, `B_NAME`, `A_INCOME`, `B_INCOME`, `DATE_NIGHTS`, `CHECKINS`,
  `SAVED_THIS_MONTH` constants and the `FAIRNESS_GOAL`, `DATE_GOAL`, `CHECKIN_GOAL`,
  `SAVINGS_GOAL`, `CHORE_RATIO_GOAL` targets plus the `BILLS`, `CHORES`, `GOALS`,
  `SAVINGS`, `DATES`, `CHECKINS_LOG`, `BIG_TALKS`, `ADMIN`, `INDIVIDUAL`, `MONTHS` tables
  in `build_xlsx.py`.
- **Keep the tie-outs**: `SAVINGS` monthly must sum to `SAVED_THIS_MONTH` ($940),
  `INDIVIDUAL`'s bill row must match `FairA` / `FairB`, and `PaysA` / `PaysB` on the Fair
  Share tab must equal the fair split if you want fairness to read 100%.
- **If you change the chore numbers, re-check the Together Score.** The 90% is built on
  housework landing at exactly 0.40 (goal 1.15 ÷ ratio 2.875).
- Keep `build_marketing.py`'s KPIs in sync with the workbook.

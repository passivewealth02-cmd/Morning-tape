# Preschool Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/preschool-command-center/build
python3 build_xlsx.py      # -> ../Preschool_Command_Center.xlsx  (17 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step quick guide + a "grows at their own
   pace" note.
2. **Dashboard** fills 12 KPI cards + a "Growing Well?" bars panel + a themes
   donut + a "coming up" list + a Ready-for-K gauge + a Skills-by-Domain bar chart.
3. **Skills & Milestones** covers 7 domains with Mastered / Emerging / Not Yet
   color-coding; the summary block feeds the dashboard bar chart.
4. **ABC & 123** tracks sees / says / writes per letter; "letters known of 26"
   updates live (**24 of 26**).
5. Marking a skill Mastered, a letter Yes, an activity Done or a readiness item
   Yes updates the dashboard + Ready Score (**70%**). Sample: Days **96**/150,
   Letters **24**/26, Skills **25**/35, Themes **7**/14, Activities **62%**,
   Ready for K **75%**, Supplies **$350**.
6. No broken cells; custom tables (Daily Rhythm, Supplies, Goals, Portfolio)
   start in column B.

> Note: uses `COUNTIF/COUNTIFS`, `AVERAGE`, `IFERROR` — opens in Google Sheets or
> Excel 2019/365. Change days goal, books or supplies budget in Settings and the
> dashboard adjusts.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Preschool_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band & gold rules. 300 DPI,
US Letter. Twelve pages: child-at-a-glance, weekly theme plan, daily rhythm,
skills & milestones checklist, ABC & 123 chart, activity planner, read-aloud log,
nature & field trips, arts & sensory bank, kindergarten-readiness checklist,
preschool goals, and a portfolio / keepsake page.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard), everything-inside (17 tabs),
the whole-child skills tracker, ABC & 123 letter-tile mastery, activity planner +
readiness, and the **12-page printables showcase** (real PDF thumbnails). Images
3–5 each show a different tab.

**Four detailed images**: 07 feature spotlights (incl. printable thumbnails),
08 "basic printable vs Command Center", 09 up-and-running in 4 steps, 10
what's-included / who-it's-for / guarantee. Ten images — fills all 10 Etsy slots.
All headline numbers (24/26 · 25/35 · 7/14 · 62% · 38 · $350 · 70%) are verified
against the workbook.

---

## D. Etsy delivery package

```
Preschool_Command_Center.xlsx     ← Google Sheets / Excel master (17 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt    ← "Make a Copy" link
Preschool_Printables.pdf           ← 12-page print-ready pack
START_HERE.pdf                     ← onboarding quick-start
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| PSC-GS   | Google Sheets only | $14 |
| PSC-PDF  | Printable PDF only | $14 |
| PSC-BUNDLE | Sheets + PDF + Quick-Start | **$22** |
| PSC-PLUS | Bundle + tot-school add-on | $28 |
| PSC-PRO  | Provider / commercial-use license | $69 |

- **Evergreen niche with a July–September peak** (new-year planning) and a
  January bump. The skills-tracker + Ready-for-K angle is far more differentiated
  than a plain preschool worksheet pack, and it fits any play-based approach.
- Two angles: **"play with a plan"** and **"Google Sheets + printables in one."**
  Bundle is the hero SKU; the tot-school add-on and provider license lift AOV.

---

## F. Maintenance

- Edit the `CHILDREN`, `SKILLS`, `ABC`, `THEMES`, `ACTIVITIES`, `READALOUD`,
  `TRIPS`, `READINESS`, `GOALS` constants in `build_xlsx.py`; every KPI + the
  Ready Score recompute. Add a child → add a `CHILDREN` row (a profile block
  renders). Skills, themes & activities are plain tables — add a row anytime.
- Printable pages live in `build_pdf.py` (one function per page) — restyle once,
  re-run. Keep `build_marketing.py` numbers in sync with the workbook.

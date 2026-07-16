# Tot-School Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/tot-school-command-center/build
python3 build_xlsx.py      # -> ../Tot_School_Command_Center.xlsx  (17 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step quick guide + a "grows on their own
   timeline" note.
2. **Dashboard** fills 12 KPI cards + a "Growing Well?" bars panel + a themes
   donut + a "coming up" list + a Growing gauge + a Milestones-by-Domain bar chart.
3. **Milestones** covers 6 domains with Met / Emerging / Not Yet color-coding;
   the summary block feeds the dashboard bar chart.
4. **First Words & Concepts** tracks colors, shapes, body parts, animal sounds &
   counting; "concepts got" updates live (**18 of 24**), plus a words counter.
5. Marking a milestone Met, a concept Yes, a tray Done or a readiness item Yes
   updates the dashboard + Growing Score (**69%**). Sample: Days **72**/120,
   Words **45**, Milestones **26**/35, Themes **5**/12, Trays **62%**,
   Preschool-ready **88%**, Supplies **$259**.
6. No broken cells; custom tables (Daily Rhythm, Supplies, Goals, Portfolio)
   start in column B.

> Note: uses `COUNTIF/COUNTIFS`, `AVERAGE`, `IFERROR` — opens in Google Sheets or
> Excel 2019/365. Change days goal, words or supplies budget in Settings and the
> dashboard adjusts.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Tot_School_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band & gold rules. 300 DPI,
US Letter. Twelve pages: tot-at-a-glance, weekly theme plan, daily rhythm,
milestones checklist, first words & concepts, tot-tray planner, board-book log,
outings & nature, sensory play bank, ready-for-preschool checklist, tot goals,
and a portfolio-of-firsts page.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard), everything-inside (17 tabs),
the toddler milestones tracker, first words & concept-tile mastery, tot-tray
planner + readiness, and the **12-page printables showcase** (real PDF
thumbnails). Images 3–5 each show a different tab.

**Four detailed images**: 07 feature spotlights (incl. printable thumbnails),
08 "basic printable vs Command Center", 09 up-and-running in 4 steps, 10
what's-included / who-it's-for / guarantee. Ten images — fills all 10 Etsy slots.
All headline numbers (26/35 · 45 words · 5/12 · 62% · 24 · $259 · 88% · 69%) are
verified against the workbook.

---

## D. Etsy delivery package

```
Tot_School_Command_Center.xlsx    ← Google Sheets / Excel master (17 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt    ← "Make a Copy" link
Tot_School_Printables.pdf          ← 12-page print-ready pack
START_HERE.pdf                     ← onboarding quick-start
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| TSC-GS   | Google Sheets only | $14 |
| TSC-PDF  | Printable PDF only | $14 |
| TSC-BUNDLE | Sheets + PDF + Quick-Start | **$22** |
| TSC-PLUS | Bundle + preschool add-on | $28 |
| TSC-PRO  | Provider / commercial-use license | $69 |

- **Evergreen niche with a strong new-parent & new-year draw.** The
  milestones + Growing Score angle is far more differentiated than a plain
  toddler activity printable, and it fits any gentle/play-based approach.
- Two angles: **"no pressure, all play"** and **"Google Sheets + printables in
  one."** Bundle is the hero SKU; the preschool add-on (upsell to the Preschool
  Command Center) and provider license lift AOV.

---

## F. Maintenance

- Edit the `TOTS`, `MILESTONES`, `CONCEPTS`, `THEMES`, `TRAYS`, `BOARDBOOKS`,
  `OUTINGS`, `READINESS`, `GOALS` constants in `build_xlsx.py`; every KPI + the
  Growing Score recompute. Add a tot → add a `TOTS` row (a profile block
  renders). Milestones, concepts & trays are plain tables — add a row anytime.
- Printable pages live in `build_pdf.py` (one function per page) — restyle once,
  re-run. Keep `build_marketing.py` numbers in sync with the workbook.

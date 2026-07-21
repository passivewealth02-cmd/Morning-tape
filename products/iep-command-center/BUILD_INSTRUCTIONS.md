# IEP Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/iep-command-center/build
python3 build_xlsx.py      # -> ../IEP_Command_Center.xlsx  (16 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a clear "not medical or
   legal advice / keep private" note.
2. **IEP Goals** takes baseline / target / current for each goal; Progress
   calculates itself. Sample goals: **65 / 64 / 50 / 76 / 63%** → avg **64%**.
3. **Dashboard** fills 12 KPI cards + a "How We're Doing" table + a Progress-
   Toward-Each-Goal bar chart. **Progress Score 80%**.
4. **Services** shows minutes scheduled vs delivered (**420 / 395**);
   **Accommodations** flags each Yes/No (**9 of 10** in place).
5. Changing a goal's Current level re-runs its progress and the dashboard.
6. No broken cells; custom tables (IEP Goals, Services, Strengths, Health) start
   in column B.

> Note: uses `IF`, `IFERROR`, `MIN/MAX`, `COUNTIF`, `AVERAGE`, `SUM` — opens in
> Google Sheets or Excel 2019/365. All goal progress is live.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../IEP_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light advocacy-binder forms on white with a forest-green header band. 300
DPI, US Letter. Twelve pages: profile & team, goals & progress, progress
monitoring, services & minutes, accommodations checklist, therapy log, behavior
tracker (ABC), meeting notes / prep, communication log, strengths & interests,
records & documents, and wins & milestones.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard), everything-inside (16 tabs), IEP
goals & progress (the trending bars), services & accommodations, data + wins, and
the **12-page advocacy-binder showcase**. Images 3–5 each show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "a paper folder vs Command
Center", 09 ready-to-advocate in 4 steps, 10 what's-included / who-it's-for /
good-to-know. Ten images — fills all 10 Etsy slots. All headline numbers (5 goals
· 64% avg · 4 on pace · 420 min · 9/10 accoms · 80% score) are verified against
the workbook.

---

## D. Etsy delivery package

```
IEP_Command_Center.xlsx            ← Google Sheets / Excel master (16 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
IEP_Printables.pdf                  ← 12-page printable advocacy binder
START_HERE.pdf                      ← onboarding quick-start
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| IEP-GS   | Google Sheets only | $15 |
| IEP-PDF  | Printable PDF only | $15 |
| IEP-BUNDLE | Sheets + PDF + Quick-Start | **$24** |
| IEP-PLUS | Bundle + sibling / multi-child add-on | $32 |
| IEP-PRO  | Advocate / commercial-use license | $79 |

- **A deeply-motivated, underserved buyer.** Parents of kids with an IEP are
  organized, invested and often overwhelmed — a calm, meeting-ready system stands
  out. Steady year-round demand with a back-to-school and annual-review bump.
- Two angles: **"walk in prepared"** and **"your child is more than a plan."**
  Keep the tone warm and by-a-parent; lead the FAQ with the "not legal advice"
  note to build trust.

---

## F. Maintenance

- Edit the `STUDENT`, `GOALS`, `MONITORING`, `SERVICES`, `ACCOMMODATIONS`,
  `THERAPY`, `BEHAVIOR`, `MEETINGS`, `COMMUNICATION`, `RECORDS`, `WINS` constants
  in `build_xlsx.py`; every KPI + the Progress Score recompute. Goals are named
  cells — change baseline/target/current and progress follows.
- Printable pages live in `build_pdf.py` (one function per page). Keep
  `build_marketing.py` numbers in sync with the workbook.

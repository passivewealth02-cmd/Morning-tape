# Home Renovation & Remodel Budget Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/reno-budget-command-center/build
python3 build_xlsx.py      # -> ../Reno_Budget_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not construction/financial
   advice" note.
2. **Rooms** sums five rooms to **TOTAL BUDGET $60,000** (`TotalBudget`) and **TOTAL
   SPENT $42,000** (`TotalSpent`), leaving **$18,000** (`Remaining`, 70% used).
   `RoomsUnder` = 5 via `SUMPRODUCT`; `RoomCount` = 5.
3. **Budget vs Actual** shows the remaining **$18,000** and a **$9,000** contingency
   reserve (`ContingencyReserve` = 15% × budget), **$7,000** left (`ContingencyLeft`).
4. **Payments** sums to **PAID TO DATE $38,000** (`PaidToDate`); **$4,000**
   outstanding (`Outstanding` = spent − paid).
5. **Change Orders** total **$3,000** (`ChangeOrders`); **% complete** **65%**
   (`PctComplete`).
6. **Dashboard** fills 12 KPI cards + a Reno Health table + budget-used & spent-by-
   month charts. **Reno Score 90%** (change-orders-in-check is the honest weak
   dimension). No broken cells; custom tables start in column B.

> Note: uses `SUM`, `SUMPRODUCT`, `SUMIF`, `COUNTA`, `COUNTIF`, `INDEX`, `AVERAGE`,
> `MIN`, `IF`, `IFERROR` — opens in Google Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Reno_Budget_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter.
Twelve pages: budget worksheet, room budget, line items, contractor list, payment log,
change orders, materials, timeline, financing plan, decisions log, monthly summary and a
remodel checklist.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with the budget-used & spent charts),
everything-inside (14 tabs), the budget-vs-actual engine, the room budget, the remodel
engine (budget + rooms), and the **12-page printables showcase**. Images 3–5 each show a
different tab.

**Four detailed images**: 07 feature spotlights, 08 "basic budget vs Command Center",
09 run-your-remodel in 4 steps, 10 what's-included / who-it's-for / works-with. Ten
images — fills all 10 Etsy slots. All headline numbers ($60k budget · $42k spent · $18k
remaining · 70% used · $9k contingency · $7k left · 5 rooms · $38k paid · $4k
outstanding · $3k change orders · 65% complete · 90% score) are verified against the
workbook.

---

## D. Etsy delivery package

```
Reno_Budget_Command_Center.xlsx    ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
Reno_Budget_Printables.pdf          ← 12-page print-ready pack
START_HERE.pdf                      ← onboarding quick-start
```

---

## E. Pricing  (all SKUs are the same digital files — a product, never a service)

| SKU | What the buyer downloads | Price |
| --- | ------ | ----- |
| RBC-GS   | The Google Sheets / Excel file only | $19 |
| RBC-PDF  | The printable PDF only | $19 |
| RBC-BUNDLE | The spreadsheet + the printable PDF | **$29** |
| RBC-COMM | The same files + a commercial-use file license | $49 |

> Etsy Services-policy note: keep every listing a plain digital download. Do not
> advertise setup help, custom work, consultations, "done-for-you" builds, or "free
> updates / lifetime access" — offering a service (rather than a finished file) is
> what gets a listing removed and earns a strike. A commercial-use license is a
> permission term attached to the *file* and is allowed; doing the work *for* the
> buyer is not.

- **Strong demand** in spring/summer (renovation season) and January (new-year
  projects). Renovation-budget templates command premium prices — a remodel is a big
  spend and the tool is cheap insurance.
- Use all 10 photos + a walkthrough video that shows the *file*. Lead photo = the
  feature-forward hero; the **budget-vs-actual engine** and the **contingency reserve**
  are your strongest differentiators — most listings are a blank budget grid.
- Cross-sell the **Net Worth & FIRE** template — same homeowner/money-minded buyer.

---

## F. Maintenance

- Edit the `ROOMS`, `LINE_ITEMS`, `CONTRACTORS`, `PAYMENTS`, `CHANGE_ORDERS`,
  `MATERIALS`, `TIMELINE`, `FINANCING`, `DECISIONS`, `MONTHS` constants and the
  `CONTINGENCY_RATE`, `CONTINGENCY_USED_N`, `PCT_COMPLETE_N`, `ROOM_GOAL`,
  `OUTSTANDING_GOAL`, `CHANGE_GOAL` targets in `build_xlsx.py`; every KPI + the Reno
  Score recompute. Everything is cross-linked — change a room budget or a change order
  and the remaining, contingency and score follow.
- Keep `build_marketing.py`'s KPIs in sync with the workbook.

# WORKFLOW — building one product end to end

Directory layout for every product:

```
products/<slug>/
├── README.md  GOOGLE_SHEETS.md  BUILD_INSTRUCTIONS.md  ETSY_LISTING.md
├── <Name>.xlsx                       (the workbook master)
├── <Name>_Printables.pdf             (dual-format only)
└── build/
    ├── build_xlsx.py
    ├── build_marketing.py
    ├── build_marketing_detail.py
    └── build_pdf.py                  (dual-format only)
└── marketing/  01..10.png  (+ print/ page PNGs for dual-format)
```

## 1 · Design

- **Name:** `<Niche> Command Center™` (e.g. "Lyft Driver Command Center™").
  Dual-format family/school products can be `<Niche> Command Center™` too.
- **Sample persona:** a believable named subject (a driver, a family, a shopper)
  so the dashboard reads like a real account.
- **Tabs (~18):** `Dashboard` first, then ~16 domain tabs, then `Settings`.
  Log-style tabs → `build_log`. Profile/budget/dashboard/summary tabs → custom.
- **12 dashboard KPIs:** pick the numbers a buyer would brag about. Always end
  with a blended **Health / Readiness / On-Track score** = `AVERAGE(HealthRange)`
  over ~6 dimension cells (each `=IFERROR(MIN(actual/target,1),0)` or a completion %).

## 2 · Lock the numbers FIRST (non-negotiable)

Write sample data as explicit Python constants (`SHIFTS`, `CHILDREN`, `SUPPLIES`,
`WEEKS`, `BUDGET`, …). Run a replica to compute every KPI and the health score:

```bash
python3 -c "from build_xlsx import SHIFTS; hrs=sum(s[2] for s in SHIFTS); ..."
```

Only continue once the numbers are clean and quotable. These exact numbers go on
the dashboard KPI cards AND in every marketing image AND in the README's
"Verified sample" line. They must all agree.

Cross-link interdependent money: e.g. the Expenses "Fuel" line = `=FuelTotal`
(named total from the log), so the log and expenses never disagree. A revenue
line that also appears elsewhere references the other cell by name.

## 3 · build_xlsx.py

- Copy `helpers_xlsx.py` verbatim (palette, styles, `luxe_header`, `build_log`,
  `kpi_card`, `nrange`, `cell_name`, …).
- `build_settings(wb)`: input cells (named) + dropdown-list banks (named).
- One builder per tab. Log tabs: `build_log(...)` then add named ranges + a TOTAL
  row + conditional formatting. Custom tabs: `set_widths([2, ...])` and write
  content starting at **column 2** (`table_headers(..., start_col=2)` and
  `ws.cell(row=r, column=2, ...)`), banners `merge_set("B..:E..")`.
- `build_dashboard(wb)`: create at index 0. Two rows of 6 `kpi_card`s at rows 5 &
  8, cols `[2,4,6,8,10,12]`, span 2. A "health" table whose score column is named
  `HealthRange`; `AVERAGE(HealthRange)` is the headline score. Add 2–4 charts
  (DoughnutChart/BarChart/LineChart) referencing other sheets; `dataLabels =
  no_labels()`.
- `main()`: build all, set `wb._sheets = [wb[n] for n in order]`, save to
  `../<Name>.xlsx`.

### Dashboard KPI card block (the pattern)

```python
row1 = [("LABEL", "=NamedOrFormula", "money|num|pct|money2|dec|text"), ...6]
row2 = [ ...6 ]
cols6 = [2,4,6,8,10,12]
for (lab,fml,kind),col in zip(row1,cols6): kpi_card(ws,5,col,2,lab,fml,kind)
for (lab,fml,kind),col in zip(row2,cols6): kpi_card(ws,8,col,2,lab,fml,kind)
```

### Health score pattern

An Analytics/Dashboard block holds ~6 dimension cells, each `=IFERROR(MIN(actual/
target,1),0)` (or a completion ratio). Name that column `HealthRange`. The
headline KPI = `=IFERROR(AVERAGE(HealthRange),0)`.

## 4 · Verify

```bash
python3 build_xlsx.py
python3 -c "import openpyxl; wb=openpyxl.load_workbook('../<Name>.xlsx'); \
print(len(wb.sheetnames)); print([n for n in NEEDED if n not in wb.defined_names] or 'OK')"
```

Confirm: sheet count, all named ranges resolve, and for each custom table
`ws['A5'].value is None` and `ws['B5'].value` is your first field (column-B rule).
LibreOffice headless recalc may be unavailable in the sandbox — the Python replica
is the source of truth for numbers.

## 5 · build_marketing.py (6 app-screenshots)

Copy `helpers_marketing.py`. Then define, in the product file:
- the **crest** function (see BRAND.md),
- `TABS` (sidebar tab names) and `FILE_LABEL` (window title bar),
- `KPIS` = the 12 verified `(LABEL, value, sub)` tuples,
- `content_dashboard(img,cbox)` — 12 KPI cards + 4 panels (bars, a donut, a small
  list, a gauge donut) using `hbars`/`donut`/`legend`,
- `content_<sheet>(img,cbox)` for the 3 sheet-spread images (use `_table`),
- `render_hero/inside/<three sheets>/mobile-or-printables`.

Image plan: **01 hero** (crest + wordmark + 3 stat chips + `app_window` dashboard
+ bottom pill), **02 inside** (grid of all tabs), **03/04/05 distinct sheets**
(never repeat the hero), **06 mobile** preview (or the printables showcase for
dual-format).

## 6 · build_marketing_detail.py (4 benefit images)

`from build_marketing import (...helpers, crest...)`. Render **07 features** (4
spotlight cards: a donut, a mini table, a stock list, a bars/print block), **08
compare** (basic-X vs Command Center check/cross table), **09 how-it-works** (4
numbered steps), **10 value** (what's included / who it's for / guarantee +
wordmark). Ten images fill all 10 Etsy photo slots.

## 7 · build_pdf.py (dual-format only)

Copy `helpers_pdf.py`. One function per page using `page/checkbox/field/section/
table`. Keep pages **ink-light** (white, green header band, gold rule). Save a
multi-page US-Letter PDF (`imgs[0].save(pdf, "PDF", save_all=True,
append_images=imgs[1:])`) and page PNGs into `marketing/print/`. Reuse those PNGs
as the **06 printables showcase** image (a 4×3 thumbnail grid).

## 8 · Docs, commit, deliver

- `README.md` (product overview, tab table, KPI table, verified-sample line,
  design notes, build commands).
- `GOOGLE_SHEETS.md` (Settings controls, cross-sheet named ranges, dashboard +
  health formulas, brand palette).
- `BUILD_INSTRUCTIONS.md` (how to build/verify, marketing, delivery package,
  pricing table, maintenance).
- `ETSY_LISTING.md` (see ETSY_TEMPLATE.md).
- `git`: `rm -rf build/__pycache__`, `git add products/<slug>`, commit to the
  **designated feature branch** with the required footer, `git push -u origin <branch>`.
- **Deliver via chat (SendUserFile):** absolute paths. Marketing images in
  batches of 3 with `display:"render"`. The `.xlsx` / PDF / `ETSY_LISTING.md`
  with `display:"attach"`.

## Errors seen before (avoid them)

- Triple-quoted formula strings — never put `"=IF(F5="""",...)"` ; use `""`.
- Writing to a merged cell after `merge_set` (read-only) — don't restyle merged cells.
- Emoji/`■`/`⭐` as tofu boxes — use DejaVu-safe glyphs or vector icons; `★`=U+2605.
- Doubled loops / off-by-one in sample builders — keep one clean loop.
- KPI mismatch between marketing and workbook — always re-run the replica.
- **Mileage-vs-actual honesty:** if actual expenses exceed the mileage deduction,
  the Tax Center auto-picks "Actual" — say so; don't claim "Mileage" wins.
- **Column-A margin bug:** custom tables must start at column 2 (see rule 5).

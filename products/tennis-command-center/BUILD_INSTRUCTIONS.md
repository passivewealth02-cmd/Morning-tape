# Tennis Command Center™ — Build Instructions

---

## A. Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/tennis-command-center/build
python3 build_xlsx.py      # -> ../Tennis_Command_Center.xlsx  (19-tab system + Welcome)
```

### Verifying
1. **Welcome** shows the intro + a Start-Here quick guide.
2. **Executive Tennis Dashboard** fills 12 KPI cards + 4 charts (match results,
   rating progress, skill progress, budget breakdown).
3. **Match Tracker** — 26 matches, **65% win rate** (17–9), Sets Won 37, Games
   Won 313; each Result derives from its set score; W/L glow mint/red.
4. **Match Analytics** ranks serve/rally stats (top rating 9.0); **Tournaments**
   show 2 titles + 3 upcoming registrations.
5. **Skills** chart start-vs-now (avg 7.2/10); **Equipment** flags 2 items due;
   **Budget** shows $1,800 of $1,975 monthly. **Analytics** blends a **Player
   Performance Score** (63%). No broken cells.

> Note: uses `SUMPRODUCT`, `COUNTIFS`, `AVERAGEIF` — opens in Excel 2019/365 or
> Google Sheets.

---

## B. Google Sheets build

Follow `GOOGLE_SHEETS.md`: build **Settings** first, define the cross-sheet
named ranges, then the Match Tracker & trackers, then the **Dashboard** +
**Analytics**.

---

## C. Marketing images

```bash
cd build && python3 build_marketing.py     # -> ../marketing/01..06.png
```

Six 2000×2000 PNGs, rendered as dense app screenshots (sidebar of all 19 tabs
+ the real computed KPI numbers + fully populated tables/charts): hero,
everything-inside (19-tab showcase), match tracker, match analytics, skills +
tournaments, and mobile. (Image 3–5 each show a different sheet — no repeat of
the hero dashboard.)

---

## D. Etsy delivery package

```
Tennis_Command_Center.xlsx        ← Excel master (19-tab system + Welcome)
GOOGLE_SHEETS_TEMPLATE_LINK.txt   ← "Make a Copy" link
START_HERE.pdf                    ← onboarding quick-start
THANK_YOU.pdf                     ← brand thank-you card
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| TNC-EX     | Excel only | $19 |
| TNC-GS     | Google Sheets only | $19 |
| TNC-BUNDLE | Excel + Google Sheets + Quick-Start PDF | **$29** |
| TNC-PLUS   | Bundle + drills & strategy playbook | $39 |
| TNC-PRO    | Coach / academy / creator license | $99 |

- **Strong evergreen niche** with a spend-happy audience (junior tennis
  parents, competitive adults). Bumps in **spring/summer** (season) and
  **January** (new-year goals).
- Lean into two angles: **"train like a pro"** (competitive juniors/adults)
  and **"organize the whole season"** (parents & coaches). Level-agnostic reach.
- Strong **coach/academy** upsell (TNC-PRO) for managing multiple athletes.

---

## F. Maintenance

- Edit sample data in `build_xlsx.py` and rerun; win %, sets/games, rating and
  the performance score recompute automatically. Match scores are entered
  set-by-set and totals derive themselves.
- Brand styles, dropdowns, KPI cards & summary tables are centralized.
- New sheet: add a builder and slot it into the `order` list in `main()`.
- Marketing numbers in `build_marketing.py` (`KPIS` list + content functions)
  should be kept in sync with the workbook's sample data.

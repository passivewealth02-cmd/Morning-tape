# Next Chapter™ — Divorce Organization & Life Rebuild System

> A calm, professional command center in Excel & Google Sheets — organize documents, finances, schedules, parenting, and the goals that rebuild your life.

| | |
| - | - |
| **Product** | Next Chapter™ |
| **Target** | People preparing for / going through / recovering from divorce · divorce coaches & mediators (client resource) · family-law professionals |
| **Angle** | Not a PDF checklist — a 20-sheet life-transition operating system. |
| **Formats** | Excel `.xlsx` (20 sheets) + Google Sheets edition |
| **Pricing** | $24 single · **$34 bundle** · $59 with onboarding · $129 coach/pro license |

> ⚠️ **Organizational & planning tool only.** Next Chapter™ is **not** legal,
> financial, or mental-health advice. Laws vary by jurisdiction. The
> disclaimer ships on the **Welcome** tab.

---

## Contents

```
products/next-chapter/
├── README.md
├── Next_Chapter.xlsx          ← Excel master (20 sheets)
├── GOOGLE_SHEETS.md
├── BUILD_INSTRUCTIONS.md
├── ETSY_LISTING.md
└── build/
    ├── build_xlsx.py
    └── build_marketing.py
```

---

## The 20-sheet system

| # | Sheet | # | Sheet |
| - | ----- | - | ----- |
| 1 | Welcome (+ disclaimer) | 11 | Parenting Organizer |
| 2 | Life Dashboard | 12 | Expense Tracker |
| 3 | Personal Profile | 13 | Appointments |
| 4 | Master Timeline | 14 | Contact Directory |
| 5 | Master Task Manager | 15 | Goal Planner |
| 6 | Document Vault Index | 16 | Life Rebuild Planner |
| 7 | Financial Snapshot | 17 | Journal & Notes |
| 8 | Monthly Budget | 18 | Document Checklist |
| 9 | Property & Asset Inventory | 19 | Analytics |
| 10 | Debt Tracker | 20 | Settings |

---

## Signature automation

| Metric | How it's computed |
| ------ | ----------------- |
| Days in Process | `=MAX(TODAY()-ProcessStart,0)` |
| Tasks Completed | `=COUNTIF(TaskStatus,"Complete")` |
| Documents Collected | `=COUNTIF(DocStatus,"Collected")` |
| Upcoming Appointments | `=SUMPRODUCT((ApptDate>=TODAY())*…)` |
| Savings Balance | `=SUMIFS(SnapValue,SnapType,"Asset",SnapCat,"Cash & Savings")` |
| Net Worth | `=TotalAssets-TotalDebts` |
| Budget Actuals | `=SUMIF(ExpCat, …, ExpAmount)` from the Expense Tracker |
| Parenting Set | `=COUNTA(ParentSchedule)/7` |
| Overall Progress | blend of task, document & goal completion |

41 named ranges keep the dashboards wired; blank-safe totals avoid broken
cells; all charts are cleanly placed (no overlaps, no jumbled labels).

---

## Calm, professional, empowering design

- Two-row **gold-divider headers** on every tab
- Gold-topped white **KPI cards**
- Status color-coding (Complete = mint, overdue = soft red, needed = gold)
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`
- Supportive, forward-looking copy — never clinical

---

## Build & ship

```bash
cd build && python3 build_xlsx.py        # -> ../Next_Chapter.xlsx
python3 build_marketing.py               # -> ../marketing/01..06.png
```

See `BUILD_INSTRUCTIONS.md` and `ETSY_LISTING.md`.

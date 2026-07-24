# Meal Planning & Grocery Budget Command Center™ — The Kitchen Operating System

> Not a meal planner — a **complete plan-it, cost-it, save-on-it system**.
> One premium **Google Sheets + printable PDF** command center for the kitchen: a
> cost-per-meal engine (recipe ingredients ÷ servings → cost per plate vs eating out),
> a weekly plan, recipes, a price book, a grocery list, a pantry, a budget, an
> eating-out log, a savings roll-up, meal ideas and a monthly summary — everything
> cross-linked and live.

| | |
| - | - |
| **Product** | Meal Planning & Grocery Budget Command Center™ |
| **Target** | Busy families & couples · budget & frugal-living fans · meal preppers & batch cooks · new cooks setting up a kitchen · anyone cutting the grocery bill |
| **Angle** | Plan the week, know the cost, and save on every meal. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $19 single file · **$29 bundle** (Sheets + PDF) · $49 with a commercial-use file license — all the *same digital files*, never a service |

---

## Contents

```
products/meal-plan-command-center/
├── README.md
├── Meal_Plan_Command_Center.xlsx     ← Google Sheets / Excel master (14 tabs)
├── Meal_Plan_Printables.pdf          ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 8 | Pantry |
| 2 | Dashboard | 9 | Budget |
| 3 | Cost Per Meal | 10 | Eating Out |
| 4 | Weekly Plan | 11 | Savings |
| 5 | Recipes | 12 | Meal Ideas |
| 6 | Price Book | 13 | Monthly Summary |
| 7 | Grocery List | 14 | Settings |

## The 12 printable PDF pages

Weekly Meal Plan · Cost Per Meal · Grocery List · Price Book · Recipe Rotation · Pantry
Inventory · Grocery Budget · Eating Out Log · Meal Ideas · Savings Summary · Monthly
Summary · Kitchen Checklist.

---

## Signature automation — know the cost of every plate

Everything connects. A recipe's ingredients set its cost, servings set the cost per
plate, and the eat-out price sets what you save — every meal:

```
Cost per serving = recipe ingredient cost ÷ servings
You save/serving = eat-out price per serving − cost per serving
Monthly savings  = servings cooked × eat-out price − grocery spent
Budget used      = spent this month ÷ grocery budget
```

### The 12 dashboard KPIs
Recipe Cost · Servings · Cost/Serving · Eat-Out/Serving · Saved/Serving · Weekly Plan ·
Grocery Budget · Spent This Month · Budget Used · Servings Cooked · Monthly Savings ·
Kitchen Score. The **Kitchen Score** blends under-budget, cooking-saves, week-planned,
cheap-per-serving, price-book-built and eating-out-rare into one 0–100% number.

**Verified sample kitchen** (Hearth & Harvest, cook Nora): recipe cost **$10.00** ·
servings **4** · cost/serving **$2.50** · eat-out/serving **$12.50** · saved/serving
**$10.00** · weekly plan **$66** · grocery budget **$600** · spent **$540** · budget
used **90%** · servings cooked **120** · monthly savings **$960** · **Kitchen Score
90%**.

---

## Premium kitchen design

- A **cost-per-meal** engine (ingredients ÷ servings → cost per plate)
- What you **save vs eating out**, on every serving and every month
- A **weekly plan** and a **recipe** rotation with cost per serving
- A store **price book**, a **grocery list** and a **pantry**
- A monthly **budget**, an **eating-out** log and a **savings** roll-up
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A budgeting & organizing tool, not financial or nutrition advice.** Confirm figures
> with your own sources.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Meal_Plan_Command_Center.xlsx
python3 build_pdf.py                         # -> ../Meal_Plan_Printables.pdf (+ page PNGs)
python3 build_marketing.py                  # -> ../marketing/01..06.png
python3 build_marketing_detail.py           # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

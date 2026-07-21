# Fitness & Meal-Prep Command Center™ — The Complete Health System

> Not a tracker — a **complete plan-the-food, log-the-lifts, track-the-body
> system**. One premium **Google Sheets + printable PDF** command center for your
> whole health routine: goals & stats, a weekly meal plan, a reusable recipe
> bank, an auto grocery list, a macro tracker, a weekly workout plan, a workout
> log, body metrics and a habit tracker — all rolling up into one live Fitness
> Score.

| | |
| - | - |
| **Product** | Fitness & Meal-Prep Command Center™ |
| **Target** | Weight-loss & recomp journeys · gym-goers who count macros · meal-preppers & batch cookers · beginners who want structure · couples getting fit together · anyone building a healthy habit |
| **Angle** | Plan the food, log the lifts, track the body & build the habit — one system. |
| **Formats** | Google Sheets (12-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $16 single · **$24 bundle** (Sheets + PDF) · $32 with the macro-coaching add-on · $79 coach / commercial license |

---

## Contents

```
products/fitness-command-center/
├── README.md
├── Fitness_Command_Center.xlsx   ← Google Sheets / Excel master (12 tabs)
├── Fitness_Printables.pdf        ← 12-page print-ready pack (US Letter)
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

## The 12-tab system

| # | Tab | # | Tab |
| - | --- | - | --- |
| 1 | Start Here | 7 | Macro Tracker |
| 2 | Dashboard | 8 | Workout Plan |
| 3 | Goals & Stats | 9 | Workout Log |
| 4 | Meal Plan | 10 | Body Metrics |
| 5 | Recipe Bank | 11 | Habit Tracker |
| 6 | Grocery List | 12 | Settings |

## The 12 printable PDF pages

Weekly Meal Planner · Grocery List (by aisle) · Macro Tracker · Recipe Cards ·
Workout Plan · Workout Log · Body Measurements · Weight Progress Chart · Habit
Tracker · Meal-Prep Day Checklist · Progress & Measurements · Goals & Why.

---

## Signature automation — one Fitness Score from your whole routine

The macro tracker averages calories & protein against your targets, the workout
plan counts sessions done vs your weekly goal, body metrics chart the weight
trend, and the habit tracker averages water, sleep & steps. It all rolls into a
live **Fitness Score**.

### The 12 dashboard KPIs
Current Weight · Goal Weight · Lbs to Go · Lost So Far · Calorie Target · Avg
Calories · Protein Target · Avg Protein · Workouts/Wk · Steps Avg · Water Avg ·
Fitness Score. The **Fitness Score** blends weight-to-goal, protein hit,
calories-on-target, workouts done, steps and water into one 0–100% number.

**Verified sample person** (Jordan): Current **176 lb** → goal **165 lb**
(**11** to go, **9 lb** lost in 8 weeks) · Calorie target **2,100** / avg
**2,062** · Protein target **150 g** / avg **149 g** · Workouts **4 / 5** this
week · Steps avg **9,493** · Water avg **7.1** cups · Total lifting volume
**37,515** · **Fitness Score 84%**.

---

## Premium fitness-software design

- A **weekly meal plan** pulling from a reusable **recipe bank** with macros
- An **auto grocery list** by aisle and a **macro tracker** vs your targets
- A **weekly workout split**, a **workout log** with total volume, and a
  **body-metrics** weight-trend line chart
- A **habit tracker** for water, sleep & steps, plus a live Fitness Score
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A general wellness tool, not medical, nutrition or training advice.** Talk to
> a doctor or qualified professional before starting any new diet or program.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Fitness_Command_Center.xlsx
python3 build_pdf.py                        # -> ../Fitness_Printables.pdf (+ page PNGs)
python3 build_marketing.py                 # -> ../marketing/01..06.png
python3 build_marketing_detail.py          # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

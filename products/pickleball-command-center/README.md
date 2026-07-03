# Pickleball Command Center™ — The Ultimate Pickleball Performance, League & Tournament Management System

> Not a score tracker — a **complete pickleball operating system**. One premium
> Excel & Google Sheets dashboard for matches, skills, tournaments, practice,
> equipment, fitness, finances, partners, travel, nutrition & goals.

| | |
| - | - |
| **Product** | Pickleball Command Center™ |
| **Target** | Beginner → tournament players · league players · seniors · coaches · doubles partners · clubs & tournament directors |
| **Angle** | Play once a week or compete every weekend — improve your game and stay organized. |
| **Formats** | Excel `.xlsx` (17-tab system + Welcome) + Google Sheets edition |
| **Pricing** | $19 single · **$29 bundle** · $39 with drills & strategy playbook · $79 coach/club license |

---

## Contents

```
products/pickleball-command-center/
├── README.md
├── Pickleball_Command_Center.xlsx    ← Excel master (17-tab system + Welcome)
├── GOOGLE_SHEETS.md
├── BUILD_INSTRUCTIONS.md
├── ETSY_LISTING.md
└── build/
    ├── build_xlsx.py
    └── build_marketing.py
```

---

## The 17-sheet system

| # | Sheet | # | Sheet |
| - | ----- | - | ----- |
| 1 | Executive Dashboard | 10 | Doubles Partner Tracker |
| 2 | Player Profile | 11 | Club & League Manager |
| 3 | Match Tracker | 12 | Travel Planner |
| 4 | Tournament Manager | 13 | Nutrition & Hydration |
| 5 | Practice Planner | 14 | Goal Tracker |
| 6 | Skill Development | 15 | Photo & Memory Gallery |
| 7 | Equipment Command Center | 16 | Analytics Dashboard |
| 8 | Pickleball Budget | 17 | Settings |
| 9 | Fitness & Recovery | | |

*(+ a Welcome / Start-Here tab — 18 tabs total.)*

---

## Signature automation (12 dashboard KPIs)

| Metric | How it's computed |
| ------ | ----------------- |
| Matches Played | `=COUNTA(MatchResult)` |
| Win Percentage | `=COUNTIF(MatchResult,"W")/COUNTA(MatchResult)` |
| Current Rating | `=PlayerRating` (DUPR-style) |
| Tournament Wins | `=COUNTIF(TournPlace,"1st")` |
| Practice (30 days) | `=COUNTIFS(PracticeDate,">="&TODAY()-30)` |
| Court Time (hrs) | `=SUM(PracticeHrs)+COUNTA(MatchResult)` |
| Monthly Budget | `=MonthlyBudget` |
| Gear to Replace | `=SUMPRODUCT((EquipReplace<=TODAY()+60)*…)` |
| Fitness Progress | `=AVERAGEIF(GoalCategory,"Fitness",GoalProgress)` |
| Upcoming Events | `=COUNTIF(TournDate,">="&TODAY())` |
| Top Partner | `INDEX(PartnerRecord, MATCH(MAX(win%)…))` → e.g. `10-4` |
| Season Progress | `=COUNTA(MatchResult)/SeasonTarget` |

Match results feed win %, partner records & the results donut; the Skills tab
charts **start-to-now progress**; equipment fires **replacement reminders**;
and a **Pickleball Performance Score** blends wins, skills, fitness,
tournaments, goals & practice consistency. **48 named ranges**, blank-safe
`IFERROR` formulas, cleanly-placed charts.

---

## Premium sports-software design

- Two-row **gold-divider headers** on every tab
- Gold-topped white **KPI cards** + a true executive dashboard (12 KPIs +
  match-results, court-time, skill-progress & spending charts)
- Wins glow mint / losses flag red; 1st-place & registered tournaments flag;
  gear "due soon" flags gold; skill ratings heat-map
- Image-placeholder **Photo Gallery** for medals, teams & big wins
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

---

## Build & ship

```bash
cd build && python3 build_xlsx.py        # -> ../Pickleball_Command_Center.xlsx
python3 build_marketing.py               # -> ../marketing/01..06.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

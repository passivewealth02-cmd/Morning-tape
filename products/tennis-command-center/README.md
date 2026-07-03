# Tennis Command Center™ — The Ultimate Tennis Performance, Training & Tournament Management System

> Not a tennis planner — a **complete tennis operating system**. One premium
> Excel & Google Sheets dashboard for matches, analytics, tournaments,
> practice, skills, fitness, equipment, finances, coaching & long-term
> development — with academy-level reporting and automation.

| | |
| - | - |
| **Product** | Tennis Command Center™ |
| **Target** | Beginner → tournament players · juniors, HS & college · adult league · coaches, academies & parents of junior players |
| **Angle** | Train like a pro, organize like an academy — measure progress across every part of your game. |
| **Formats** | Excel `.xlsx` (19-tab system + Welcome) + Google Sheets edition |
| **Pricing** | $19 single · **$29 bundle** · $39 with drills & strategy playbook · $99 coach/academy license |

---

## Contents

```
products/tennis-command-center/
├── README.md
├── Tennis_Command_Center.xlsx       ← Excel master (19-tab system + Welcome)
├── GOOGLE_SHEETS.md
├── BUILD_INSTRUCTIONS.md
├── ETSY_LISTING.md
└── build/
    ├── build_xlsx.py
    └── build_marketing.py
```

---

## The 19-sheet system

| # | Sheet | # | Sheet |
| - | ----- | - | ----- |
| 1 | Executive Tennis Dashboard | 11 | Doubles Partner Tracker |
| 2 | Player Profile | 12 | Coaching Notes |
| 3 | Match Tracker | 13 | Goal Command Center |
| 4 | Match Analytics | 14 | Travel Planner |
| 5 | Tournament Command Center | 15 | Nutrition & Recovery |
| 6 | Practice Planner | 16 | Photo & Video Library |
| 7 | Skill Development Center | 17 | Season Planner |
| 8 | Fitness & Conditioning | 18 | Analytics Command Center |
| 9 | Equipment Command Center | 19 | Settings |
| 10 | Tennis Budget | | |

*(+ a Welcome / Start-Here tab — 20 tabs total.)*

---

## Signature automation (12 dashboard KPIs)

| Metric | How it's computed |
| ------ | ----------------- |
| Matches Played | `=COUNTA(MatchResult)` |
| Win Percentage | `=COUNTIF(MatchResult,"W")/COUNTA(MatchResult)` |
| Current Ranking | `="#"&CurrentRank` |
| Current Rating | `=PlayerRating` (UTR-style) |
| Practice Hours | `=SUM(PracticeHrs)` |
| Tournament Wins | `=COUNTIF(TournPlace,"1st")` |
| Sets Won | `=SUM(MatchSetsW)` |
| Games Won | `=SUM(MatchGamesW)` |
| Fitness Score | `=AVERAGEIF(GoalCategory,"Fitness",GoalProgress)` |
| Gear to Replace | `=SUMPRODUCT((EquipReplace<=TODAY()+45)*…)` |
| Monthly Budget | `=MonthlyBudget` |
| Player Performance Score | `=AVERAGE(Analytics!C7:C12)` |

Match results auto-compute **win %, sets, games & surface records** (a match's
Result is derived from its set score); **Match Analytics** tracks serve, rally
& break-point stats with match ratings; equipment fires **replacement
reminders**; and a **Player Performance Score** blends wins, skills, fitness,
tournaments, goals & practice. **52 named ranges**, blank-safe `IFERROR`
formulas, cleanly-placed charts.

---

## Premium academy-software design

- Two-row **gold-divider headers** on every tab
- Gold-topped white **KPI cards** + a true executive dashboard (12 KPIs +
  match-results, rating-progress, skill & spending charts)
- Wins glow mint / losses flag red; 1st-place & registered events flag;
  serve-% heat-maps; gear "due soon" flags gold; a periodized **Season Planner**
- Image-placeholder **Photo & Video Library** for technique review
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

---

## Build & ship

```bash
cd build && python3 build_xlsx.py        # -> ../Tennis_Command_Center.xlsx
python3 build_marketing.py               # -> ../marketing/01..06.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

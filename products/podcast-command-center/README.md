# Podcast Command Center™ — The Ultimate Podcast Production & Business System

> Not an episode tracker — a **complete Podcast Operating System**. One premium
> Excel & Google Sheets command center for episodes, guests, recording, show
> notes, downloads analytics, platforms, sponsors, memberships, clips and
> finances — everything that turns listens into real income.

| | |
| - | - |
| **Product** | Podcast Command Center™ |
| **Target** | Independent podcasters · interview & solo shows · video podcasters · podcast networks & producers · coaches & educators · agencies managing shows |
| **Angle** | Publish consistently, grow downloads, land sponsors — run the business behind your show. |
| **Formats** | Excel `.xlsx` (24-tab system + Welcome) + Google Sheets edition |
| **Pricing** | $19 single · **$29 bundle** · $39 with launch & sponsorship playbook · $99 network / producer license |

---

## Contents

```
products/podcast-command-center/
├── README.md
├── Podcast_Command_Center.xlsx      ← Excel master (24-tab system + Welcome)
├── GOOGLE_SHEETS.md
├── BUILD_INSTRUCTIONS.md
├── ETSY_LISTING.md
└── build/
    ├── build_xlsx.py
    ├── build_marketing.py
    └── build_marketing_detail.py
```

---

## The 24-sheet system

| # | Sheet | # | Sheet |
| - | ----- | - | ----- |
| 1 | Show Dashboard | 13 | Clips & Promo |
| 2 | Show Profile | 14 | Repurposing Engine |
| 3 | Episode Calendar | 15 | Finance Center |
| 4 | Production Pipeline | 16 | Expenses |
| 5 | Episode Planner | 17 | Equipment |
| 6 | Guest CRM | 18 | Reviews & Ratings |
| 7 | Recording Log | 19 | Brand Kit |
| 8 | Show Notes & Assets | 20 | Cover & Clip Gallery |
| 9 | Analytics Center | 21 | Goals & OKRs |
| 10 | Platform Analytics | 22 | Audience Insights |
| 11 | Sponsors & Ads | 23 | Collabs & Cross-Promo |
| 12 | Memberships | 24 | Settings |

*(+ a Welcome / Start-Here tab — 25 tabs total.)*

---

## Signature automation (12 dashboard KPIs)

| Metric | How it's computed |
| ------ | ----------------- |
| Subscribers | `=Subscribers` |
| Downloads (28d) | `=Downloads28` |
| Avg / Episode | `=AvgPerEp` |
| Episodes (28d) | `=COUNTIFS(CalDate…"Published")` |
| Consumption | `=Consumption` |
| Monthly Revenue | `=RevenueTotal` (7 streams) |
| Net Profit | `=RevenueTotal-ExpenseTotal` |
| Ad Rate (CPM) | `=CPM` |
| Active Sponsors | `=COUNTIF(SponStage,"Booked")+COUNTIF(SponStage,"Live")` |
| Members | `=MemberCount` |
| Publishing Consistency | `=MIN(Episodes28/EpGoal,1)` |
| Show Health Score | `=AVERAGE(HealthRange)` |

Membership tiers roll into **members & monthly revenue** that flow straight into
the Finance Center; every income & expense becomes **live monthly revenue,
run-rate and net profit**; the episode calendar's status calculates itself; the
sponsor pipeline tracks **lead → paid**; and a **Show Health Score** blends
revenue, publishing, downloads, consumption, sponsor pipeline and audience growth.

**Verified sample show** (Make It Work, an 18.5k-subscriber business & creativity
podcast): Subscribers **18.5K** · Downloads **42K/28d** (11K/episode) ·
Consumption **78%** · Revenue **$9,130/mo** across 7 streams · Net **$7,230**
(79% margin) · Members **260** · Show Health **83%**.

---

## Premium producer-software design

- Two-row **gold-divider headers** on every tab
- Gold-topped white **KPI cards** + a true show dashboard (12 KPIs +
  downloads-trend, revenue-by-source, top-episodes & expense charts)
- Published episodes glow, sponsor stages & guest statuses color-code; an
  interest score ranks episode ideas; downloads data-bars on every episode
- Image-placeholder **Cover & Clip Gallery** (Insert ▸ Picture-in-cell or
  `=IMAGE()`)
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Podcast_Command_Center.xlsx
python3 build_marketing.py                 # -> ../marketing/01..06.png
python3 build_marketing_detail.py          # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

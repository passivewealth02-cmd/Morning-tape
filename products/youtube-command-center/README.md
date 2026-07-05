# YouTube Command Center™ — The Ultimate YouTube Business Operating System

> Not a content planner — a **complete YouTube business operating system**. One
> premium Excel & Google Sheets dashboard for content, analytics, SEO,
> sponsorships, affiliates, products, finances and long-term growth — with
> media-company-level reporting and automation.

| | |
| - | - |
| **Product** | YouTube Command Center™ |
| **Target** | New → full-time creators · faceless channels · educators & coaches · agencies managing creators |
| **Angle** | Run your channel like a media company — measure and monetize every part of the business. |
| **Formats** | Excel `.xlsx` (30-tab system + Welcome) + Google Sheets edition |
| **Pricing** | $19 single · **$29 bundle** · $39 with creator playbook · $99 agency/creator license |

---

## Contents

```
products/youtube-command-center/
├── README.md
├── YouTube_Command_Center.xlsx      ← Excel master (30-tab system + Welcome)
├── GOOGLE_SHEETS.md
├── BUILD_INSTRUCTIONS.md
├── ETSY_LISTING.md
└── build/
    ├── build_xlsx.py
    └── build_marketing.py
```

---

## The 30-sheet system

| # | Sheet | # | Sheet |
| - | ----- | - | ----- |
| 1 | Executive YouTube Dashboard | 16 | Affiliate Tracker |
| 2 | Channel Profile | 17 | Digital Products |
| 3 | Content Master Calendar | 18 | Business Finance Center |
| 4 | Video Production Pipeline | 19 | Equipment Manager |
| 5 | Idea Vault | 20 | Repurposing Engine |
| 6 | Script Manager | 21 | Brand Kit |
| 7 | Thumbnail Lab | 22 | Asset Library |
| 8 | SEO & Keywords | 23 | AI Prompt Library |
| 9 | Analytics Command Center | 24 | Goals & OKRs |
| 10 | Long-Form Analytics | 25 | Collab Tracker |
| 11 | Shorts Dashboard | 26 | Audience Insights |
| 12 | Community Tab | 27 | Tax Center |
| 13 | Live Streams | 28 | Annual Plan |
| 14 | Playlists | 29 | Content Gallery |
| 15 | Sponsorship CRM | 30 | Settings |

*(+ a Welcome / Start-Here tab — 31 tabs total.)*

---

## Signature automation (12 dashboard KPIs)

| Metric | How it's computed |
| ------ | ----------------- |
| Subscribers | `=SubNow` |
| Views (28d) | `=Views28` |
| Watch Hours | `=WatchHrs` |
| Published (28d) | `=COUNTIFS(CalDate,">="&TODAY()-28,CalStatus,"Published")` |
| Monthly Revenue | `=RevenueTotal` (8 income streams) |
| Net Profit | `=RevenueTotal-ExpenseTotal` |
| RPM | `=RPM` |
| Avg CTR | `=AvgCTR` |
| Avg View Duration | `=AvgViewDur` |
| Brand Deals | `=COUNTIF(SponStage,"Signed")+…+"Negotiation")` |
| Upload Consistency | `=MIN(Published28/UploadGoal,1)` |
| Channel Health | `=AVERAGE(HealthRange)` |

Every income & expense rolls into **live monthly revenue, annual run-rate and
net profit**; the calendar's publishing status calculates itself; the
Sponsorship CRM tracks **lead → paid**; and a **Channel Health Score** blends
revenue-vs-goal, upload consistency, CTR, retention, sponsorship pipeline and
goal progress. Dozens of **named ranges**, blank-safe `IFERROR` formulas, and
cleanly-placed charts throughout.

**Verified sample channel** ("Vale Studio", 84.2k subs): Monthly Revenue
**$9,840** across 8 streams · Net Profit **$7,430** (75% margin) · Upload
Consistency **86%** · Brand Deals **4** active · Channel Health **79%**.

---

## Premium media-company design

- Two-row **gold-divider headers** on every tab
- Gold-topped white **KPI cards** + a true executive dashboard (12 KPIs +
  subscriber-growth, revenue-by-source, top-videos & expense charts)
- Published videos glow mint, scheduled flag gold; a **Channel Health Score**
  heat-maps its dimensions; an **AI Opportunity Score** ranks ideas
- Image-placeholder **Content Gallery** + **Thumbnail Lab** for A/B testing
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

---

## Build & ship

```bash
cd build && python3 build_xlsx.py        # -> ../YouTube_Command_Center.xlsx
python3 build_marketing.py               # -> ../marketing/01..06.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

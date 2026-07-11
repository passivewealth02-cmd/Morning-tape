# Instagram Command Center™ — The Ultimate Instagram Creator Business System

> Not a content planner — a **complete Instagram Operating System**. One premium
> Excel & Google Sheets command center for ideas, captions, grid, reels, stories,
> analytics, community, IG Shop, affiliate, brand deals and finances —
> everything that turns saves into real income.

| | |
| - | - |
| **Product** | Instagram Command Center™ |
| **Target** | Instagram creators & influencers · UGC creators · IG Shop sellers · niche / aesthetic pages · coaches & educators · agencies managing creators |
| **Angle** | Post a cohesive feed, grow reach, and run the business behind your content. |
| **Formats** | Excel `.xlsx` (24-tab system + Welcome) + Google Sheets edition |
| **Pricing** | $19 single · **$29 bundle** · $39 with creator growth playbook · $99 agency / creator license |

---

## Contents

```
products/instagram-command-center/
├── README.md
├── Instagram_Command_Center.xlsx    ← Excel master (24-tab system + Welcome)
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
| 1 | Creator Dashboard | 13 | Instagram Shop |
| 2 | Creator Profile | 14 | Affiliate Tracker |
| 3 | Content Calendar | 15 | Brand Deals (UGC CRM) |
| 4 | Grid Planner | 16 | Finance Center |
| 5 | Content Pipeline | 17 | Expenses |
| 6 | Idea Vault | 18 | Equipment |
| 7 | Captions & Hooks | 19 | Repurposing Engine |
| 8 | Reels Analytics | 20 | Brand Kit |
| 9 | Stories Planner | 21 | Content Gallery |
| 10 | Hashtags & SEO | 22 | Goals & OKRs |
| 11 | Analytics Center | 23 | Audience Insights |
| 12 | Community & DMs | 24 | Settings |

*(+ a Welcome / Start-Here tab — 25 tabs total.)*

---

## Signature automation (12 dashboard KPIs)

| Metric | How it's computed |
| ------ | ----------------- |
| Followers | `=FollowerNow` |
| Reach (28d) | `=Reach28` |
| Engagement Rate | `=EngRate` |
| Saves (28d) | `=Saves28` |
| Posted (28d) | `=COUNTIFS(CalDate…"Posted")+COUNTIFS(…"Boosted")` |
| Monthly Revenue | `=RevenueTotal` (7 streams) |
| Net Profit | `=RevenueTotal-ExpenseTotal` |
| Shop Sales (GMV) | `=ShopGMV` |
| Brand Deals | `=COUNTIF(SponStage,"Signed")+…+"Negotiation")` |
| Follower Growth | `=FollowerGrowth` |
| Posting Consistency | `=MIN(Posted28/PostGoal,1)` |
| Creator Health Score | `=AVERAGE(HealthRange)` |

IG Shop units roll into **GMV and commission**, which flow straight into the
Finance Center; every income & expense becomes **live monthly revenue, run-rate
and net profit**; the calendar's posting status calculates itself; the UGC CRM
tracks **lead → paid**; and a **Creator Health Score** blends revenue, posting,
engagement, reel retention, brand pipeline and follower growth. The **Grid
Planner** previews your next 9 posts for a cohesive feed.

**Verified sample account** (@studiofern, a 148k-follower plant & home creator):
Followers **148K** · Reach **1.15M/28d** · Engagement **6.4%** · Saves **42K** ·
Revenue **$8,069/mo** across 7 streams · Net **$6,419** (80% margin) · Shop GMV
**$9,758** · Creator Health **87%**.

---

## Premium creator-software design

- Two-row **gold-divider headers** on every tab
- Gold-topped white **KPI cards** + a true creator dashboard (12 KPIs +
  follower-growth, revenue-by-source, top-posts & expense charts)
- Boosted posts glow, deal stages color-code, story links flag; an **AI
  Opportunity Score** ranks ideas and a **caption bank** scores what earns saves
- Image-placeholder **Grid Planner** (3×3 feed preview) + **Content Gallery**
  (Insert ▸ Picture-in-cell or `=IMAGE()`)
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Instagram_Command_Center.xlsx
python3 build_marketing.py                 # -> ../marketing/01..06.png
python3 build_marketing_detail.py          # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

# TikTok Command Center™ — The Ultimate TikTok Creator Business System

> Not a content calendar — a **complete TikTok Operating System**. One premium
> Excel & Google Sheets command center for ideas, hooks, trends, analytics,
> LIVE, TikTok Shop, affiliate, brand deals and finances — everything that
> turns views into real income.

| | |
| - | - |
| **Product** | TikTok Command Center™ |
| **Target** | TikTok creators & influencers · UGC creators · TikTok Shop sellers · faceless / niche pages · coaches & educators · agencies managing creators |
| **Angle** | Post more consistently, ride trends earlier, and run the business behind your content. |
| **Formats** | Excel `.xlsx` (24-tab system + Welcome) + Google Sheets edition |
| **Pricing** | $19 single · **$29 bundle** · $39 with creator growth playbook · $99 agency / creator license |

---

## Contents

```
products/tiktok-command-center/
├── README.md
├── TikTok_Command_Center.xlsx       ← Excel master (24-tab system + Welcome)
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
| 1 | Creator Dashboard | 13 | Brand Deals (UGC CRM) |
| 2 | Creator Profile | 14 | Series & Playlists |
| 3 | Content Calendar | 15 | Finance Center |
| 4 | Video Pipeline | 16 | Expenses |
| 5 | Idea Vault | 17 | Equipment |
| 6 | Hook & Script Bank | 18 | Repurposing Engine |
| 7 | Trends & Sounds | 19 | Brand Kit |
| 8 | Hashtag & SEO | 20 | Content Gallery |
| 9 | Analytics Center | 21 | Goals & OKRs |
| 10 | LIVE Sessions | 22 | Audience Insights |
| 11 | TikTok Shop | 23 | Collabs & Duets |
| 12 | Affiliate Tracker | 24 | Settings |

*(+ a Welcome / Start-Here tab — 25 tabs total.)*

---

## Signature automation (12 dashboard KPIs)

| Metric | How it's computed |
| ------ | ----------------- |
| Followers | `=FollowerNow` |
| Views (28d) | `=Views28` |
| Engagement Rate | `=EngRate` |
| Avg Completion | `=Completion` |
| Posted (28d) | `=COUNTIFS(CalDate…"Posted")+COUNTIFS(…"Viral")` |
| Monthly Revenue | `=RevenueTotal` (7 streams) |
| Net Profit | `=RevenueTotal-ExpenseTotal` |
| Shop Sales (GMV) | `=ShopGMV` |
| Brand Deals | `=COUNTIF(SponStage,"Signed")+…+"Negotiation")` |
| Follower Growth | `=FollowerGrowth` |
| Posting Consistency | `=MIN(Posted28/PostGoal,1)` |
| Creator Health Score | `=AVERAGE(HealthRange)` |

TikTok Shop units roll into **GMV and commission**, which flow straight into the
Finance Center; every income & expense becomes **live monthly revenue, run-rate
and net profit**; the calendar's posting status calculates itself; the UGC CRM
tracks **lead → paid**; and a **Creator Health Score** blends revenue, posting,
engagement, retention, brand pipeline and follower growth.

**Verified sample account** (@vellacreates, a 215k-follower DIY & lifestyle
creator): Followers **215K** · Views **3.2M/28d** · Engagement **9.2%** ·
Revenue **$8,583/mo** across 7 streams · Net **$6,633** (77% margin) · Shop GMV
**$14,233** · Creator Health **88%**.

---

## Premium creator-software design

- Two-row **gold-divider headers** on every tab
- Gold-topped white **KPI cards** + a true creator dashboard (12 KPIs +
  follower-growth, revenue-by-source, top-videos & expense charts)
- Viral posts glow, trends flag Rising/Peaking/Fading, deal stages color-code;
  an **AI Opportunity Score** ranks ideas and a **hook bank** scores what works
- Image-placeholder **Content Gallery** for covers (Insert ▸ Picture-in-cell or
  `=IMAGE()`)
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../TikTok_Command_Center.xlsx
python3 build_marketing.py                 # -> ../marketing/01..06.png
python3 build_marketing_detail.py          # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

# Personal & Private Chef Command Center™ — The Private-Chef Operating System

> Not a price list — a **complete cost-it, price-it, pay-yourself system**.
> One premium **Google Sheets + printable PDF** command center for a personal or
> private chef: a per-event pricing engine (cost a dinner per guest, set your
> per-guest price, and see your take-home and real hourly rate), a service menu &
> dish costing, a client roster that turns bookings into monthly revenue, a booking
> calendar, grocery & kitchen-kit lists, a mileage log, waste, income & expenses and
> a monthly summary — everything cross-linked and live.

| | |
| - | - |
| **Product** | Personal & Private Chef Command Center™ |
| **Target** | Personal & private chefs · in-home dinner-party chefs · meal-prep & drop-off chefs · yacht & villa chefs · caterers going solo · anyone cooking for private clients |
| **Angle** | Cost every dinner, and pay yourself properly. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $19 single file · **$29 bundle** (Sheets + PDF) · $49 with a commercial-use file license — all the *same digital files*, never a service |

---

## Contents

```
products/private-chef-command-center/
├── README.md
├── Private_Chef_Command_Center.xlsx  ← Google Sheets / Excel master (14 tabs)
├── Private_Chef_Printables.pdf       ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 8 | Grocery List |
| 2 | Dashboard | 9 | Kitchen Kit |
| 3 | Event Pricing | 10 | Mileage & Travel |
| 4 | Service Menu | 11 | Waste Log |
| 5 | Dish Costing | 12 | Income & Expenses |
| 6 | Clients | 13 | Monthly Summary |
| 7 | Booking Calendar | 14 | Settings |

## The 12 printable PDF pages

Event Quote · Service Menu · Dish Cost Card · Client Roster · Prep List · Shopping
List · Kitchen Kit · Event Run Sheet · Mileage & Travel · Income & Expenses · Monthly
Summary · Event Day Checklist.

---

## Signature automation — cost the dinner, pay yourself

Everything connects. Your course costs build the food per guest, your per-guest price
sets the event price, and your hours turn it into a real hourly rate:

```
Food per guest   = Σ (appetizer + main + sides + dessert + pantry)
Event food       = Food per guest × Guests
Event price      = Price per guest × Guests
Your take-home   = Event price − Event food − Travel
Real hourly rate = Take-home ÷ Hours worked
Monthly revenue  = Σ (client events/month × rate)
```

### The 12 dashboard KPIs
Food/Guest · Price/Guest · Food Cost % · Margin/Event · Your Hourly · Top Client ·
Monthly Revenue · Active Clients · Events/Month · Monthly Profit · Waste % · Chef
Score. The **Chef Score** blends food-cost-on-target, margin-healthy, menu-priced,
clients-vs-goal, profitable and waste-low into one 0–100% number.

**Verified sample chef** (Chef's Table Co., owner Rowan): food **$22.00**/guest ·
price **$90**/guest · food cost **24%** · margin **70%**/event · your hourly
**$54.00** · top client **Anderson** ($2,160/mo) · monthly revenue **$5,950** · **5**
clients · **13** events/month · monthly profit **$3,340** · waste **2.4%** · **Chef
Score 90%**.

---

## Premium private-chef design

- A **per-event pricing** engine (cost per guest → take-home & real hourly rate)
- A **service menu** and **dish costing** to quote fast and stay profitable
- A **client roster** that turns bookings into **monthly revenue**
- A **booking calendar**, **grocery & kitchen-kit** lists and a **mileage log**
- **Waste**, **income & expenses** and a **monthly summary** that builds by season
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A business tool, not financial, legal or food-safety advice.** Follow your local
> food-handling rules and confirm figures with your own advisors.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Private_Chef_Command_Center.xlsx
python3 build_pdf.py                         # -> ../Private_Chef_Printables.pdf (+ page PNGs)
python3 build_marketing.py                  # -> ../marketing/01..06.png
python3 build_marketing_detail.py           # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

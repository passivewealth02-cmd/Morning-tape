# Meal Prep Business Command Center™ — The Meal-Prep Operating System

> Not a menu — a **complete cost-it, price-it, grow-the-MRR system**.
> One premium **Google Sheets + printable PDF** command center for a meal-prep
> business: a per-meal cost engine (protein, carb, veg, sauce, packaging & labor),
> meal-plan pricing at 5/10/15/21 meals a week, a subscriber board that turns counts
> into weekly revenue and monthly recurring revenue, a production plan, ingredient &
> packaging costs, weekly orders & delivery, waste, income & expenses and a monthly
> summary — everything cross-linked and live.

| | |
| - | - |
| **Product** | Meal Prep Business Command Center™ |
| **Target** | Meal-prep & meal-delivery brands · subscription meal makers · personal chefs · fitness / macro meal sellers · ghost & commissary kitchens · anyone selling prepped meals |
| **Angle** | Cost every meal, and grow your recurring revenue. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $19 single file · **$29 bundle** (Sheets + PDF) · $49 with a commercial-use file license — all the *same digital files*, never a service |

---

## Contents

```
products/meal-prep-command-center/
├── README.md
├── Meal_Prep_Command_Center.xlsx     ← Google Sheets / Excel master (14 tabs)
├── Meal_Prep_Printables.pdf          ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 8 | Ingredient Costs |
| 2 | Dashboard | 9 | Packaging & Delivery |
| 3 | Meal Cost | 10 | Weekly Orders |
| 4 | Meal Plans | 11 | Waste Log |
| 5 | Menu | 12 | Income & Expenses |
| 6 | Subscribers | 13 | Monthly Summary |
| 7 | Production Plan | 14 | Settings |

## The 12 printable PDF pages

Meal Cost Card · Meal Plan Pricing · Menu · Subscriber List · Prep List · Ingredient
Cost List · Packaging & Delivery · Delivery Run Sheet · Waste Log · Income & Expenses
· Monthly Summary · Prep Day Checklist.

---

## Signature automation — cost per meal, and MRR

Everything connects. Your meal cost feeds every plan price, your subscriber counts
feed weekly revenue, and weekly revenue becomes monthly recurring revenue:

```
Cost per meal   = Σ (protein + carb + veg + sauce + packaging + labor)
Plan cost       = Meals per plan × Cost per meal
Weekly revenue  = Σ (subscribers per plan × plan price)
MRR             = Weekly revenue × 4
Food cost %     = Ingredient cost ÷ Avg price per meal
Margin per meal = (Avg price per meal − Cost per meal) ÷ Avg price per meal
```

### The 12 dashboard KPIs
Meal Cost · Avg Price/Meal · Food Cost % · Margin/Meal · Subscribers · Top Plan ·
Weekly Revenue · MRR · Meals/Week · Monthly Profit · Waste % · Prep Score. The
**Prep Score** blends food-cost-on-target, margin-healthy, plans-costed,
subscribers-vs-goal, profitable and waste-low into one 0–100% number.

**Verified sample business** (Fresh Fuel Co., owner Kai): meal cost **$5.00** · avg
price **$10.33**/meal · food cost **31%** · margin **52%**/meal · **46** subscribers
· top plan **10 meals / week** ($1,575/wk) · weekly revenue **$4,475** · **MRR
$17,900** · **433** meals/week · monthly profit **$6,740** · waste **2.4%** ·
**Prep Score 90%**.

---

## Premium meal-prep design

- A **per-meal cost** engine (protein, carb, veg, sauce, packaging & labor)
- **Meal-plan pricing** at 5 / 10 / 15 / 21 meals a week, with live margin
- A **subscriber board** that turns counts into weekly revenue & **MRR**
- **Production plan**, **ingredient & packaging costs**, **weekly orders & delivery**
- **Waste**, **income & expenses** and a **monthly summary** that compounds
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A business tool, not financial, legal or food-safety advice.** Follow your local
> food-handling rules and confirm figures with your own advisors.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Meal_Prep_Command_Center.xlsx
python3 build_pdf.py                         # -> ../Meal_Prep_Printables.pdf (+ page PNGs)
python3 build_marketing.py                  # -> ../marketing/01..06.png
python3 build_marketing_detail.py           # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

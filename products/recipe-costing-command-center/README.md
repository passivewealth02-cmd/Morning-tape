# Recipe Costing & Menu Engineering Command Center™ — The Complete Menu-Profit System

> Not a costing sheet — a **complete cost-price-engineer profit system**. One
> premium **Google Sheets + printable PDF** command center for a more profitable
> menu: an ingredient price library, a recipe-costing engine, menu items with
> live food-cost % & margin, a star/plowhorse/puzzle/dog **menu-engineering
> matrix**, a target-margin price calculator, sales mix, portion & yield,
> specials, batch scaling, a vendor price log and a waste log.

| | |
| - | - |
| **Product** | Recipe Costing & Menu Engineering Command Center™ |
| **Target** | Restaurant & café owners · chefs & kitchen managers · food trucks & caterers · bakeries & bars · ghost kitchens & pop-ups · anyone pricing a menu |
| **Angle** | Cost every plate, price with intent & engineer a more profitable menu. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $19 single · **$29 bundle** (Sheets + PDF) · $39 with the food-cost add-on · $99 multi-location / commercial license |

---

## Contents

```
products/recipe-costing-command-center/
├── README.md
├── Recipe_Costing_Command_Center.xlsx   ← Google Sheets / Excel master (14 tabs)
├── Recipe_Costing_Printables.pdf        ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 8 | Sales Mix |
| 2 | Dashboard | 9 | Portion & Yield |
| 3 | Ingredients | 10 | Specials & LTO |
| 4 | Recipe Costing | 11 | Batch & Prep |
| 5 | Menu Items | 12 | Vendor Prices |
| 6 | Menu Engineering | 13 | Waste Log |
| 7 | Price Calculator | 14 | Settings |

## The 12 printable PDF pages

Recipe Cost Card · Prep List · Menu Engineering Worksheet · Food-Cost Pricing
Guide · Ingredient Price Log · Portion & Yield Worksheet · Menu Item P&L ·
Specials & LTO Planner · Batch Recipe Scaler · Vendor Price Tracker · Waste Log ·
Weekly Food-Cost Tracker.

---

## Signature automation — cost the plate, then engineer the menu

The ingredient library turns pack price ÷ pack size into a **cost per unit**;
recipes cost themselves line by line (`qty × cost/unit`). Each menu item then
shows **food-cost % = plate cost ÷ price** and **margin = price − plate cost**,
and is auto-classified against the menu's own average popularity and margin:

```
Star       high popularity + high profit  → feature & protect
Plowhorse  high popularity + low profit    → trim cost / raise price
Puzzle     low popularity + high profit    → reposition & upsell
Dog        low popularity + low profit     → rework or cut
```

### The 12 dashboard KPIs
Menu Items · Avg Food Cost · Target Food Cost · Avg Plate Cost · Avg Menu Price ·
Avg Margin · Stars · Plowhorses · Puzzles · Dogs · Top Margin · Menu Score. The
**Menu Score** blends food-cost-on-target, margin-vs-goal, fully-costed,
above-margin-goal, menu-balance and contribution into one 0–100% number.

**Verified sample menu** (The Copper Skillet, chef-owner Marco): **8** items ·
avg food cost **24%** (6 pts under the 30% target) · avg plate cost **$4.75** ·
avg price **$18.00** · avg margin **$13.25** · **1 Star / 4 Plowhorses / 2
Puzzles / 1 Dog** · top-margin item **Ribeye Steak** ($21.50/plate) · monthly
revenue **$30,430** / profit **$22,917** · **Menu Score 92%**.

---

## Premium menu-profit-software design

- An **ingredient library** (pack price → cost per unit) and a **recipe-costing** engine
- **Menu items** with live food-cost % + margin, and the **engineering matrix**
- A **price calculator** (plate cost ÷ target food-cost %), **sales mix** & **yield**
- **Specials, batch, vendor-price** and **waste** tools that keep real costs honest
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A business tool, not financial or accounting advice.** Confirm figures with
> your own books.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Recipe_Costing_Command_Center.xlsx
python3 build_pdf.py                        # -> ../Recipe_Costing_Printables.pdf (+ page PNGs)
python3 build_marketing.py                 # -> ../marketing/01..06.png
python3 build_marketing_detail.py          # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

# Bar & Pub Command Center™ — The Complete Bar Business System

> Not a sales log — a **complete cost-the-pour, kill-the-shrinkage system**.
> One premium **Google Sheets + printable PDF** command center for a bar or pub: a
> pour-cost engine, a drink menu with pour-cost % and margin, keg & draft profit,
> inventory variance, liquor count & par, a happy-hour margin planner, waste &
> spill, weekly sales, ordering, cash & tips and events & big tabs.

| | |
| - | - |
| **Product** | Bar & Pub Command Center™ |
| **Target** | Bars, pubs & taprooms · cocktail & wine bars · breweries & beer gardens · restaurant bar programs · nightclubs & lounges · anyone pouring for profit |
| **Angle** | Cost every pour, kill your shrinkage, and pour more profit. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $19 single · **$29 bundle** (Sheets + PDF) · $39 with the pour-cost add-on · $99 commercial / multi-location license |

---

## Contents

```
products/bar-command-center/
├── README.md
├── Bar_Command_Center.xlsx           ← Google Sheets / Excel master (14 tabs)
├── Bar_Printables.pdf                ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 8 | Happy Hour |
| 2 | Dashboard | 9 | Weekly Sales |
| 3 | Pour Cost | 10 | Waste & Spill |
| 4 | Drink Menu | 11 | Ordering |
| 5 | Keg & Draft | 12 | Cash & Tips |
| 6 | Inventory Variance | 13 | Events & Tabs |
| 7 | Liquor Inventory | 14 | Settings |

## The 12 printable PDF pages

Pour Cost Card · Drink Menu · Keg & Draft Log · Inventory Variance · Liquor Count
Sheet · Happy Hour Planner · Weekly Sales Log · Waste & Spill Log · Ordering Sheet
· Cash & Tips · Events & Big Tabs · Open/Close Checklist.

---

## Signature automation — cost the pour, kill the shrinkage

The pour-cost engine costs a drink to the pour — spirit, mixer & garnish — then
each menu drink carries a pour-cost % and margin; kegs roll up to a profit per
keg, and inventory variance measures theoretical usage against the actual count:

```
Pour cost           = SUM(component qty × cost/unit)
Pour-cost %         = Pour cost ÷ Menu price
Margin per drink    = Menu price − Pour cost
Profit per keg      = (Price/pint − Cost/pint) × Pints per keg
Variance %          = (Actual usage − Theoretical usage) ÷ Theoretical usage
```

### The 12 dashboard KPIs
Drinks · Avg Pour Cost · Avg Price · Avg Margin · Pour Cost % · Top Seller ·
Weekly Sales · Weekly Units · Avg Profit/Keg · Inventory Variance · Happy-Hour
Margin · Bar Score. The **Bar Score** blends pour-cost-on-target, margin-per-drink,
menu-fully-costed, low-variance, happy-hour-margin and gross-margin into one
0–100% number.

**Verified sample bar** (The Oak & Iron, owner Reed): **8** drinks · avg pour cost
**$1.61** · avg price **$9.62** · avg margin **$8.02** · pour cost **16%** · top
seller **Draft Beer** ($2,940/wk) · weekly sales **$16,940** · **1,870** units ·
avg profit/keg **$571** · inventory variance **6.1%** · happy-hour margin **79%** ·
**Bar Score 90%**.

---

## Premium bar-software design

- A **pour-cost** engine (spirit + mixer + garnish = true pour cost)
- A **drink menu** with pour-cost % and margin on every drink
- A **keg & draft** profit tracker (cost per pint → profit per keg)
- **Inventory variance**, **liquor count & par**, a **happy-hour margin planner**
- **Waste & spill**, **weekly sales**, **ordering**, **cash & tips** & **events**
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A business tool, not financial or accounting advice.** Confirm figures with
> your own books.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Bar_Command_Center.xlsx
python3 build_pdf.py                        # -> ../Bar_Printables.pdf (+ page PNGs)
python3 build_marketing.py                 # -> ../marketing/01..06.png
python3 build_marketing_detail.py          # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

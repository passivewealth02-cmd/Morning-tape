# Cafe & Coffee Shop Command Center™ — The Complete Coffee-Shop System

> Not a sales tracker — a **complete cost-every-cup, watch-your-prime-cost
> system**. One premium **Google Sheets + printable PDF** command center for a
> café: a cup-cost engine, a menu board with beverage-cost %, daypart sales,
> weekly sales, labor & prime cost, bean & milk usage, inventory & par, a waste
> log, ordering, cash & tips and a regulars tracker.

| | |
| - | - |
| **Product** | Cafe & Coffee Shop Command Center™ |
| **Target** | Café & coffee-shop owners · baristas & shift leads · espresso bars & roasteries · drive-thru & kiosk coffee · bubble tea & juice bars · new cafés dialing in prime cost |
| **Angle** | Cost every cup, watch your prime cost & pour more profit. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $19 single · **$29 bundle** (Sheets + PDF) · $39 with the recipe-costing add-on · $99 multi-location / commercial license |

---

## Contents

```
products/cafe-command-center/
├── README.md
├── Cafe_Command_Center.xlsx          ← Google Sheets / Excel master (14 tabs)
├── Cafe_Printables.pdf               ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 8 | Bean & Milk |
| 2 | Dashboard | 9 | Inventory & Par |
| 3 | Cup Cost | 10 | Waste Log |
| 4 | Menu Board | 11 | Ordering |
| 5 | Daypart Sales | 12 | Cash & Tips |
| 6 | Weekly Sales | 13 | Regulars |
| 7 | Labor & Prime | 14 | Settings |

## The 12 printable PDF pages

Cup Cost Card · Menu Board · Daypart Sales Log · Weekly Sales Log · Labor & Prime
Cost · Bean & Milk Usage · Inventory & Par · Waste Log · Ordering Sheet · Cash &
Tips · Open / Close Checklist · Regulars & Loyalty.

---

## Signature automation — cost the cup, watch the prime cost

The cup-cost engine costs each drink to the cup (beans, milk, cup, lid, sleeve);
each menu item shows **beverage-cost % = cup cost ÷ price** and margin; and the
café's **prime cost** — the number that makes or breaks the business — is the sum
of beverage % and labor %:

```
Cup cost           = Σ (component qty × cost per unit)
Beverage cost %    = Σ(cost×units) ÷ Σ(price×units)   (weighted by mix)
Prime cost         = Beverage % + Labor %             (keep under 60%)
```

### The 12 dashboard KPIs
Menu Items · Avg Cup Cost · Avg Price · Avg Margin · Bev Cost · Labor Cost · Daily
Sales · Transactions · Avg Ticket · Top Daypart · Prime Cost · Café Score. The
**Café Score** blends beverage-cost, labor, prime cost, avg ticket, margin per
cup and low waste into one 0–100% number.

**Verified sample café** (Wildroot Coffee, owner Priya): **10** items · avg cup
cost **$1.13** · avg price **$5.05** · avg margin **$3.92** · beverage cost **22%**
· labor **28%** · daily sales **$2,769** · **352** transactions · avg ticket
**$7.87** · top daypart **Morning rush** ($1,329) · **prime cost 50%** · **Café
Score 91%**.

---

## Premium coffee-shop-software design

- A **cup-cost engine** and a **menu board** with beverage-cost % on every drink
- **Daypart** and **weekly** sales, and a **labor & prime cost** view
- **Bean & milk usage**, **inventory & par**, a **waste log** and an **ordering** sheet
- **Cash & tips** reconciliation and a **regulars** / loyalty tracker
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A business tool, not financial or accounting advice.** Confirm figures with
> your own books.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Cafe_Command_Center.xlsx
python3 build_pdf.py                        # -> ../Cafe_Printables.pdf (+ page PNGs)
python3 build_marketing.py                 # -> ../marketing/01..06.png
python3 build_marketing_detail.py          # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

# Catering Command Center™ — The Complete Catering Business System

> Not a price list — a **complete cost-the-head, quote-with-confidence system**.
> One premium **Google Sheets + printable PDF** command center for a caterer: a
> per-head plate-costing engine, a menu-package price list with margins, an event
> quote that becomes a full event P&L, staffing & labor, rentals, a bookings
> calendar, inventory, waste, ordering, cash & deposits and a client CRM.

| | |
| - | - |
| **Product** | Catering Command Center™ |
| **Target** | Full-service caterers · private & personal chefs · wedding & event caterers · corporate & drop-off catering · food trucks adding catering · anyone catering for profit |
| **Angle** | Cost every head, quote with confidence, and book more profit. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $19 single · **$29 bundle** (Sheets + PDF) · $39 with the plate-costing add-on · $99 commercial / multi-location license |

---

## Contents

```
products/catering-command-center/
├── README.md
├── Catering_Command_Center.xlsx      ← Google Sheets / Excel master (14 tabs)
├── Catering_Printables.pdf           ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 8 | Bookings |
| 2 | Dashboard | 9 | Inventory |
| 3 | Plate Costing | 10 | Waste Log |
| 4 | Menu Packages | 11 | Ordering |
| 5 | Event Quotes | 12 | Cash & Deposits |
| 6 | Staffing | 13 | Clients |
| 7 | Rentals | 14 | Settings |

## The 12 printable PDF pages

Plate Cost Card · Menu Package Price List · Event Quote Sheet · Event Run Sheet ·
Staffing Sheet · Rentals & Equipment · Bookings Calendar · Inventory & Par ·
Waste Log · Ordering Sheet · Cash & Deposits · Client Contact Sheet.

---

## Signature automation — cost the head, quote a full P&L

The plate engine costs a plate by the head; each package then carries a margin
and food-cost %. An event quote multiplies guests by the package price and layers
in service, staff and rentals to produce a full event P&L:

```
Cost per head      = SUM(component per-head costs)
Margin per head    = Price per head − Cost per head
Event revenue      = Guests × Price/head + Service fee
Event food cost    = Guests × Cost/head
Event margin       = Revenue − Food − Staff − Rentals
Event margin %     = Event margin ÷ Revenue
```

### The 12 dashboard KPIs
Events · Avg Guests · Revenue · Avg Per Head · Food Cost · Top Package · Avg Event
· Avg Margin · Labor · Packages · Waste % · Catering Score. The **Catering Score**
blends food-cost-on-target, margin-per-event, packages-fully-costed,
labor-under-control, bookings-vs-goal and gross-margin into one 0–100% number.

**Verified sample caterer** (Wildflower & Oak, owner Camille): **6** events · avg
guests **73** · revenue **$21,560** · avg per head **$49** · food cost **26%** ·
top package **Wedding Premium** ($7,550/event) · avg event value **$3,593** · avg
event margin **32%** · labor **25%** · **8** packages · waste **2.0%** ·
**Catering Score 90%**.

---

## Premium catering-software design

- A **plate cost-per-head** engine (protein + sides + dessert + overhead)
- A **menu-package** price list with margin & food-cost % on every package
- An **event quote** that becomes a full **event P&L** — profit before you say yes
- **Staffing & labor**, **rentals**, a **bookings calendar** & deposits
- **Inventory**, **waste**, **ordering**, **cash & deposits** & a **client CRM**
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A business tool, not financial or accounting advice.** Confirm figures with
> your own books.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Catering_Command_Center.xlsx
python3 build_pdf.py                        # -> ../Catering_Printables.pdf (+ page PNGs)
python3 build_marketing.py                 # -> ../marketing/01..06.png
python3 build_marketing_detail.py          # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

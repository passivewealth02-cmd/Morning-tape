# Ghost Kitchen Command Center™ — The Delivery-Only Business System

> Not a menu — a **complete beat-the-apps, know-your-real-margin system**.
> One premium **Google Sheets + printable PDF** command center for a delivery-only
> / ghost kitchen: an item-margin engine that shows the true net after the app's
> commission, a menu with net margin per item, a full P&L per platform, virtual
> brands, packaging, order volume, inventory, waste, ordering, payouts and promos.

| | |
| - | - |
| **Product** | Ghost Kitchen Command Center™ |
| **Target** | Ghost & cloud kitchens · delivery-only restaurants · virtual-brand operators · restaurants adding delivery · home cooks going delivery · anyone selling on the apps |
| **Angle** | Beat the apps — know your real margin after commission. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $19 single · **$29 bundle** (Sheets + PDF) · $39 with the margin add-on · $99 commercial / multi-location license |

---

## Contents

```
products/ghost-kitchen-command-center/
├── README.md
├── Ghost_Kitchen_Command_Center.xlsx  ← Google Sheets / Excel master (14 tabs)
├── Ghost_Kitchen_Printables.pdf       ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 8 | Order Volume |
| 2 | Dashboard | 9 | Inventory |
| 3 | Item Margin | 10 | Waste Log |
| 4 | Menu & Margins | 11 | Ordering |
| 5 | Platform P&L | 12 | Payouts |
| 6 | Virtual Brands | 13 | Promotions |
| 7 | Packaging | 14 | Settings |

## The 12 printable PDF pages

Item Margin Card · Menu & Margins · Platform P&L · Virtual Brands · Packaging
Sheet · Order Volume Log · Prep List · Inventory & Par · Waste Log · Ordering
Sheet · Payouts Sheet · Promotions Sheet.

---

## Signature automation — the margin the menu price hides

Delivery apps take 15–30% of every order, so the menu price is not your margin.
The item-margin engine starts from the app price and subtracts the commission,
food and packaging to show the true net you keep; each platform then rolls up to a
full P&L:

```
Net margin (item)  = App price × (1 − commission) − food − packaging
Net margin %       = Net margin ÷ App price
Platform gross     = Orders × Avg order value
Platform net       = Gross × (1 − commission)
Blended commission = Total commission ÷ Total gross
```

### The 12 dashboard KPIs
Menu Items · Avg App Price · Food Cost · Avg Net Margin · Blended Commission · Top
Item · Weekly Orders · Weekly Revenue · Net Payout · Avg Order · Virtual Brands ·
Kitchen Score. The **Kitchen Score** blends food-cost-on-target, net-margin-healthy,
menu-fully-costed, commission-in-control, direct-order-mix and gross-margin into
one 0–100% number.

**Verified sample kitchen** (Midnight Kitchen Collective, owner Devin): **8** menu
items · avg app price **$11.83** · food cost **26%** · avg net margin **44%** ·
blended commission **24%** · top item **Birria Tacos** ($6.11 net) · weekly orders
**800** · weekly revenue **$17,360** · net payout **$13,117** · avg order **$21.70**
· **4** virtual brands · **Kitchen Score 90%**.

---

## Premium delivery-software design

- An **item-margin** engine (app price − commission − food − packaging = net)
- A **menu** with net margin $ and % on every item
- A **full P&L per platform** (DoorDash, Uber Eats, Grubhub, Direct)
- **Virtual brands** from one kitchen, **packaging** cost & **order volume**
- **Inventory**, **waste**, **ordering**, **payouts** & a **promo ROI** tracker
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A business tool, not financial or accounting advice.** Confirm figures with
> your own books.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Ghost_Kitchen_Command_Center.xlsx
python3 build_pdf.py                        # -> ../Ghost_Kitchen_Printables.pdf (+ page PNGs)
python3 build_marketing.py                 # -> ../marketing/01..06.png
python3 build_marketing_detail.py          # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

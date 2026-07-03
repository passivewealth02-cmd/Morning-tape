# Trading Card Collection Command Center™ — The Ultimate Trading Card Collection, Value & Inventory System

> Not an inventory sheet — a **complete collection management system**. One
> premium Excel & Google Sheets dashboard for your collection, values,
> purchases, sales, trades, grading, wishlist, duplicates, decks & card photos.
> Works for Pokémon, Magic, Yu-Gi-Oh!, sports cards and every other game.

| | |
| - | - |
| **Product** | Trading Card Collection Command Center™ |
| **Target** | Collectors · competitive TCG players · investors · parents managing kids' collections · card shop owners · grading enthusiasts |
| **Angle** | Your collection is an asset — know what it's worth, always. |
| **Formats** | Excel `.xlsx` (13-tab system + Welcome) + Google Sheets edition |
| **Pricing** | $19 single · **$29 bundle** · $39 with collector's playbook · $79 shop/creator license |

---

## Contents

```
products/trading-card-command-center/
├── README.md
├── Trading_Card_Command_Center.xlsx   ← Excel master (13-tab system + Welcome)
├── GOOGLE_SHEETS.md
├── BUILD_INSTRUCTIONS.md
├── ETSY_LISTING.md
└── build/
    ├── build_xlsx.py
    └── build_marketing.py
```

---

## The 13-sheet system

| # | Sheet | # | Sheet |
| - | ----- | - | ----- |
| 1 | Collection Dashboard | 8 | Wishlist |
| 2 | Master Collection | 9 | Duplicates Manager |
| 3 | Purchase Tracker | 10 | Tournament Deck Builder |
| 4 | Sales Tracker | 11 | **Card Image Vault** (photo uploads) |
| 5 | Trade Tracker | 12 | Analytics Dashboard |
| 6 | Grading Center | 13 | Settings |
| 7 | Collection Value Analytics | | |

*(+ a Welcome / Start-Here tab — 14 tabs total.)*

---

## Signature automation (12 dashboard KPIs)

| Metric | How it's computed |
| ------ | ----------------- |
| Total Cards | `=SUM(ColQty)` |
| Collection Value | `=SUM(ColTotal)` (qty × est. value per row) |
| Purchase Cost | `=SUM(ColPaid)` |
| Profit / Loss | `=SUM(ColTotal)-SUM(ColPaid)` |
| Growth % | `=SUM(ColTotal)/SUM(ColPaid)-1` |
| Top Card Value | `=MAX(ColEach)` |
| Photos on File | `=COUNTIF(ColPhoto,"Yes")/COUNTA(ColName)` |
| Cards Graded | `=COUNTIF(GradeStatus,"Returned")` |
| Awaiting Grading | `=COUNTIF(GradeStatus,"Pending")` |
| Trades Completed | `=COUNTA(TradePartner)` |
| Wishlist Progress | purchased ÷ total wishlist |
| Collection Health | `=AVERAGE(Analytics!C7:C12)` |

The Master Collection computes **live profit/loss per card**; the Grading
Center tracks **value added by every slab** (+$997 in the sample); trades net
out value-in vs value-out; and a **Collection Health Score** blends growth,
grading, wishlist, storage, duplicates & tournament play. **45 named ranges**,
blank-safe `IFERROR` formulas, cleanly-placed charts.

---

## 📷 Card Image Vault — upload photos of your cards

- Dedicated **Card Image Vault** tab with 8 photo slots (Crown Jewel, Grail #2,
  Newest Slab, Binder Spread…) each with card, grade & value fields
- **On-sheet upload guide**: Excel → Insert ▸ Pictures ▸ This Device (365:
  *Place in Cell*); Google Sheets → Insert ▸ Image ▸ in cell, or `=IMAGE("link")`
- A **Photo?** column in the Master Collection + a **Photos on File** dashboard
  KPI track your visual coverage (67% in the sample)

---

## Premium collection-software design

- Two-row **gold-divider headers** on every tab
- Gold-topped white **KPI cards** + a true executive dashboard (12 KPIs +
  rarity, set, growth & grading charts)
- Card values **data-bar** by worth; grails flag gold; returned slabs & owned
  wishlist items glow mint; unplanned duplicates flag amber
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

---

## Build & ship

```bash
cd build && python3 build_xlsx.py        # -> ../Trading_Card_Command_Center.xlsx
python3 build_marketing.py               # -> ../marketing/01..06.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

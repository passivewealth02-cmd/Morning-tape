# Amazon FBA & Online Seller Profit Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/fba-seller-command-center/build
python3 build_xlsx.py      # -> ../FBA_Seller_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not financial/tax advice" note
   and a warning that Amazon's fees change.
2. **Profit Calculator**: $29.99 price → referral $4.50 (15%) + FBA $5.40 + storage
   $0.35 = **TOTAL FEES $10.25** (`TotalFees`); COGS $7.50 + inbound $1.20 = **LANDED
   COST $8.70** (`LandedCost`).
3. **NET PER UNIT $11.04** (`NetPerUnit`), **NET MARGIN 36.8%** (`NetMargin`), **ROI
   126.9%** (`ROI`). × 420 units = **MONTHLY PROFIT $4,637** on **$12,596** revenue.
4. **Product Catalog** recalculates net/unit per SKU using the shared `ReferralRate`,
   `FBAFee`, `StorageFee` and `InboundShip` names; negative nets go red.
5. **Inventory** computes days of cover per SKU; `FlagshipCover` = **60**. **PPC & ACoS**
   totals **20.0%** (`ACoS`); **Reviews** averages **4** per SKU (`AvgReviews`).
6. **Dashboard** fills 12 KPI cards + a Seller Health table + where-the-$29.99-goes &
   net-profit-by-month charts. **Seller Score 90%** (reviews-per-SKU is the honest weak
   dimension). No broken cells; custom tables start in column B.

> Note: uses `SUM`, `SUMIF`, `SUMPRODUCT`, `COUNTA`, `COUNTIF`, `AVERAGE`, `MIN`,
> `ROUND`, `IF`, `IFERROR` — opens in Google Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../FBA_Seller_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter.
Twelve pages: profit per unit, product catalog, fee worksheet, inventory & reorder, PPC
& ACoS log, sales log, supplier & PO log, returns log, reviews tracker, product-research
scorer, monthly summary and a launch checklist.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with the where-the-$29.99-goes & profit
charts), everything-inside (14 tabs), the true-profit engine, the product catalog, the
seller engine (both), and the **12-page printables showcase**. Images 3–5 each show a
different tab.

**Four detailed images**: 07 feature spotlights, 08 "free FBA calculator vs Command
Center", 09 run-your-brand in 4 steps, 10 what's-included / who-it's-for / works-with.
Ten images — fills all 10 Etsy slots. All headline numbers ($29.99 price · $10.25 fees ·
$8.70 landed · $11.04 net · 36.8% margin · 126.9% ROI · 420 units · $12,596 revenue ·
$4,637 profit · 20.0% ACoS · 60 days cover · 90% score) are verified against the
workbook.

---

## D. Etsy delivery package

```
FBA_Seller_Command_Center.xlsx     ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
FBA_Seller_Printables.pdf           ← 12-page print-ready pack
START_HERE.pdf                      ← onboarding quick-start
```

---

## E. Pricing  (all SKUs are the same digital files — a product, never a service)

| SKU | What the buyer downloads | Price |
| --- | ------ | ----- |
| FBA-GS   | The Google Sheets / Excel file only | $22 |
| FBA-PDF  | The printable PDF only | $22 |
| FBA-BUNDLE | The spreadsheet + the printable PDF | **$32** |
| FBA-COMM | The same files + a commercial-use file license | $55 |

> Etsy Services-policy note: keep every listing a plain digital download. Do not
> advertise setup help, custom work, consultations, **coaching or mentoring**,
> "done-for-you" builds, or "free updates / lifetime access" — offering a service
> (rather than a finished file) is what gets a listing removed and earns a strike. This
> matters here because the FBA space is full of coaching offers; keep yours a file.

- **Business buyers with real money on the line.** An FBA seller mis-costing one SKU
  loses hundreds a month — a $32 tool is trivial. Price at the top of the consumer range.
- **Steady year-round demand** with a Q4 spike (sellers costing holiday inventory) and a
  January new-seller surge.
- Use all 10 photos + a walkthrough video that shows the *file*. Lead photo = the
  feature-forward hero; the **where-the-$29.99-goes donut** is the single most
  persuasive image you have — it shows Amazon taking $10.25 of a $29.99 sale.
- Cross-sell the **Bookkeeping & Tax**, **Etsy Seller Profit** and **Maker Pricing**
  templates — same product-seller buyer, and a natural "Seller Bundle".

---

## F. Maintenance

- Edit the `SALE_PRICE`, `REFERRAL_RATE`, `FBA_FEE`, `STORAGE_FEE`, `UNIT_COGS`,
  `INBOUND_SHIP`, `UNITS_MONTH`, `AD_SPEND`, `AD_SALES`, `ON_HAND`, `DAILY_VELOCITY`,
  `AVG_REVIEWS` constants and the `MARGIN_GOAL`, `ROI_GOAL`, `ACOS_GOAL`, `COVER_GOAL`,
  `REVIEW_GOAL`, `SKU_GOAL` targets plus the `PRODUCTS`, `FEES`, `INVENTORY`, `SALES`,
  `PPC`, `RETURNS`, `SUPPLIERS`, `REVIEWS`, `EXPENSES`, `MONTHS` tables in
  `build_xlsx.py`; every KPI + the Seller Score recompute.
- **Re-check `FBA_FEE` and `STORAGE_FEE` against Amazon's current schedule** before each
  relist — stale fees date the screenshots and invite refund requests.
- Keep `build_marketing.py`'s KPIs in sync with the workbook.

# Handmade & Maker Pricing Command Center™ — The Maker's Operating System

> Not a price calculator — a **complete price-for-profit, pay-yourself system**.
> One premium **Google Sheets + printable PDF** command center for handmade sellers:
> a price-for-profit engine (materials + labor + overhead → wholesale & retail + your
> true hourly wage), a product line, supplies, orders, overhead & fees, a sales log,
> time & labor, expenses, reviews, channels and a monthly summary — everything
> cross-linked and live.

| | |
| - | - |
| **Product** | Handmade & Maker Pricing Command Center™ |
| **Target** | Candle & soap makers · jewelry & accessory makers · ceramic, resin & woodworkers · knit, crochet & textile makers · new & part-time makers · anyone selling handmade |
| **Angle** | Price for real profit, and finally pay yourself for your time. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $19 single file · **$29 bundle** (Sheets + PDF) · $49 with a commercial-use file license — all the *same digital files*, never a service |

---

## Contents

```
products/maker-pricing-command-center/
├── README.md
├── Maker_Pricing_Command_Center.xlsx   ← Google Sheets / Excel master (14 tabs)
├── Maker_Pricing_Printables.pdf        ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 8 | Sales Log |
| 2 | Dashboard | 9 | Time & Labor |
| 3 | Pricing Calculator | 10 | Expenses |
| 4 | Product Line | 11 | Reviews & Repeat |
| 5 | Supplies & Inventory | 12 | Channels & Markets |
| 6 | Orders | 13 | Monthly Summary |
| 7 | Overhead & Fees | 14 | Settings |

## The 12 printable PDF pages

Pricing Worksheet · Product Price List · Order Form · Supply List · Overhead & Fees ·
Sales Log · Time & Labor · Expense Log · Channel Tracker · Monthly Summary · Craft
Show Checklist · Making Checklist.

---

## Signature automation — price for profit

Everything connects. Your materials, labor and overhead set the base cost, your markup
sets the retail and wholesale price, and sales roll up to revenue and profit:

```
Base cost      = materials + (labor hours × your rate) + overhead per item
Retail price   = base cost × markup
Wholesale      = retail × wholesale factor
True hourly    = (retail − materials − overhead) ÷ labor hours
Monthly profit = revenue − materials − overhead − your paid labor
```

### The 12 dashboard KPIs
Materials · Your Labor · Base Cost · Wholesale · Retail · Retail Margin · Profit/Item ·
Your Hourly · Monthly Revenue · Units Sold · Monthly Profit · Maker Score.
The **Maker Score** blends retail-margin, markup, paying-yourself, products-priced,
profit and wholesale-margin into one 0–100% number.

**Verified sample shop** (Ember & Oak, owner Wren, flagship Soy Candle 8oz): materials
**$6.00** · labor **$5.00** · base cost **$12.00** · wholesale **$15.00** · retail
**$30.00** · retail margin **60%** · profit/item **$18.00** · your hourly **$92.00** ·
monthly revenue **$2,700** · units **90** · monthly profit **$1,350** · **Maker Score
90%**.

---

## Premium maker design

- A **price-for-profit** engine (materials + labor + overhead → wholesale & retail)
- Your **true hourly wage** on every product
- A **product line** with base cost, wholesale, retail and margin on each item
- **Supplies**, **orders**, **overhead & fees** and a **sales log**
- **Time & labor**, **expenses**, **reviews & repeat** and **channels & markets**
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A business tool, not financial, legal or tax advice.** Confirm figures with your
> own advisors.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Maker_Pricing_Command_Center.xlsx
python3 build_pdf.py                         # -> ../Maker_Pricing_Printables.pdf (+ page PNGs)
python3 build_marketing.py                  # -> ../marketing/01..06.png
python3 build_marketing_detail.py           # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

# Home Bakery & Cottage Food Command Center™ — The Home-Baker's System

> Not a price list — a **complete cost-it, price-it, pay-yourself system**.
> One premium **Google Sheets + printable PDF** command center for a home &
> cottage bakery: a "price it right" engine that pays you for your time, recipe
> costing, a product list, custom orders, ingredient costs, cottage-food labels &
> allergens, markets, income & expenses, waste, customers and a monthly summary.

| | |
| - | - |
| **Product** | Home Bakery & Cottage Food Command Center™ |
| **Target** | Home & cottage bakers · custom cake & cookie makers · farmers-market sellers · cottage-food businesses · side-hustle & hobby bakers · anyone selling what they bake |
| **Angle** | Cost every recipe, price for profit, and finally pay yourself for your time. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $19 single · **$29 bundle** (Sheets + PDF) · $39 with the pricing add-on · $99 commercial license |

---

## Contents

```
products/home-bakery-command-center/
├── README.md
├── Home_Bakery_Command_Center.xlsx     ← Google Sheets / Excel master (14 tabs)
├── Home_Bakery_Printables.pdf          ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 8 | Markets & Events |
| 2 | Dashboard | 9 | Income & Expenses |
| 3 | Price It Right | 10 | Waste Log |
| 4 | Recipe Costing | 11 | Customers |
| 5 | Product List | 12 | Monthly Summary |
| 6 | Custom Orders | 13 | Settings |
| 7 | Ingredient Costs | 14 | Labeling & Allergens |

## The 12 printable PDF pages

Price It Right · Recipe Cost Card · Product Price List · Custom Order Form ·
Ingredient Cost List · Cottage Food Label · Market Day Checklist · Income &
Expenses · Waste Log · Customer List · Monthly Summary · Order Calendar.

---

## Signature automation — pay yourself for your time

Home bakers undercharge because they forget to pay themselves. The "price it right"
engine adds ingredients, packaging, **your** labor and overhead, then shows your
profit and your true hourly wage:

```
Your labor      = (Minutes ÷ 60) × Your hourly rate
True cost       = Ingredients + Packaging + Your labor + Overhead
Profit          = Price − True cost
Effective wage  = (Price − Ingredients − Packaging − Overhead) ÷ (Minutes ÷ 60)
```

### The 12 dashboard KPIs
Products · Avg Price · Food Cost · Avg Margin · Monthly Income · Top Seller ·
Monthly Profit · Your Hourly · Open Orders · Order Value · Monthly Units · Bakery
Score. The **Bakery Score** blends food-cost-on-target, healthy-margins,
products-priced, paying-yourself, profitable and low-waste into one 0–100% number.

**Verified sample bakery** (Sugar & Thyme, owner Mel): **8** products · avg price
**$28.25** · food cost **18%** · avg margin **82%** · monthly income **$2,908** ·
top seller **Custom Cookies** ($660/mo) · monthly profit **$1,658** · your hourly
**$32.67** · **5** open orders ($416) · **113** units · **Bakery Score 90%**.

---

## Premium home-bakery design

- A **price-it-right** engine that pays you for your time (true hourly wage)
- **Recipe costing** by the batch and a **product list** with margins
- A **custom-order** tracker with deposits & due dates
- **Cottage-food labels** with allergens, **ingredient costs** & **markets**
- **Income & expenses** for tax time, **waste**, **customers** & a monthly summary
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A business tool, not financial, legal or cottage-food-law advice.** Check your
> state's cottage food rules and confirm figures with your own advisors.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Home_Bakery_Command_Center.xlsx
python3 build_pdf.py                        # -> ../Home_Bakery_Printables.pdf (+ page PNGs)
python3 build_marketing.py                 # -> ../marketing/01..06.png
python3 build_marketing_detail.py          # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

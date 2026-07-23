# Photography Business Command Center™ — The Photographer's Operating System

> Not a price list — a **complete price-it-right, book-it, profit system**.
> One premium **Google Sheets + printable PDF** command center for photographers: a
> CODB / break-even engine (what you must charge to pay yourself), a per-shoot P&L,
> packages & pricing, bookings, clients & leads, an editing queue, gear, expenses,
> mileage, reviews and a monthly summary — everything cross-linked and live.

| | |
| - | - |
| **Product** | Photography Business Command Center™ |
| **Target** | Wedding & portrait photographers · family & newborn · brand & product · event & real-estate · new & part-time photographers · anyone selling photography |
| **Angle** | Price to pay yourself, and know the profit on every shoot. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $19 single file · **$29 bundle** (Sheets + PDF) · $49 with a commercial-use file license — all the *same digital files*, never a service |

---

## Contents

```
products/photography-command-center/
├── README.md
├── Photography_Command_Center.xlsx    ← Google Sheets / Excel master (14 tabs)
├── Photography_Printables.pdf         ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 8 | Editing Queue |
| 2 | Dashboard | 9 | Gear & Inventory |
| 3 | CODB & Break-Even | 10 | Expenses |
| 4 | Shoot P&L | 11 | Mileage & Travel |
| 5 | Packages & Pricing | 12 | Reviews & Referrals |
| 6 | Bookings | 13 | Monthly Summary |
| 7 | Clients & Leads | 14 | Settings |

## The 12 printable PDF pages

CODB Worksheet · Shoot P&L · Package Menu · Bookings Calendar · Clients & Leads ·
Editing Queue · Shot List · Gear Inventory · Expense Log · Mileage Log · Monthly
Summary · Session Checklist.

---

## Signature automation — price to pay yourself

Everything connects. Your overhead and salary set the break-even price, your package
price minus costs gives the profit, and bookings roll up to revenue and net:

```
CODB / shoot   = (annual overhead + desired salary) ÷ target shoots
Net per shoot  = Package price − shoot costs
Effective rate = Net per shoot ÷ hours (shoot + edit)
Revenue (YTD)  = Σ monthly revenue
Net profit     = Revenue − business expenses
```

### The 12 dashboard KPIs
Package Price · Shoot Costs · Net/Shoot · Shoot Margin · CODB/Shoot · Effective Rate ·
Bookings YTD · Revenue YTD · Avg Booking · Net Profit · Booked Ahead · Studio Score.
The **Studio Score** blends shoot-margin, beats-break-even, effective-rate,
bookings-pace, profit and booked-ahead into one 0–100% number.

**Verified sample studio** (Amberlight Photo, owner Robin): package **$3,000** ·
costs **$800** · net/shoot **$2,200** · margin **73%** · CODB/shoot **$1,500** ·
effective rate **$88.00**/hr · **30** bookings · revenue **$45,000** · avg booking
**$1,500** · net profit **$21,000** · booked ahead **$9,600** · **Studio Score 90%**.

---

## Premium photography design

- A **CODB / break-even** engine (your minimum price per shoot)
- A **per-shoot P&L** with net and your real effective hourly rate
- **Packages & pricing** with the rate per hour on every package
- **Bookings**, **clients & leads**, an **editing queue** and **gear**
- **Expenses**, a **mileage** log, **reviews** and a **monthly summary**
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A business tool, not financial, legal or tax advice.** Confirm figures with your
> own advisors.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Photography_Command_Center.xlsx
python3 build_pdf.py                         # -> ../Photography_Printables.pdf (+ page PNGs)
python3 build_marketing.py                  # -> ../marketing/01..06.png
python3 build_marketing_detail.py           # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

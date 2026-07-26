# Small Business Bookkeeping & Tax Command Center™ — The Books & Tax Operating System

> Not a spreadsheet — a **complete track-it, total-it, file-it system**.
> One premium **Google Sheets + printable PDF** command center for your books: a
> Schedule C engine (revenue − COGS − expenses → net profit, then SE tax + income tax →
> your quarterly payment), income, expenses, COGS, mileage, sales tax, quarterly taxes,
> Schedule C line mapping, invoices, reconciliation and a monthly summary — everything
> cross-linked and live.

| | |
| - | - |
| **Product** | Small Business Bookkeeping & Tax Command Center™ |
| **Target** | Self-employed & sole proprietors · product & handmade sellers · service businesses & contractors · side hustles filing Schedule C · anyone quitting QuickBooks |
| **Angle** | Know your profit, and never be surprised by a tax bill again. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $22 single file · **$35 bundle** (Sheets + PDF) · $59 with a commercial-use file license — all the *same digital files*, never a service |

---

## Contents

```
products/bookkeeping-tax-command-center/
├── README.md
├── Bookkeeping_Tax_Command_Center.xlsx   ← Google Sheets / Excel master (14 tabs)
├── Bookkeeping_Tax_Printables.pdf         ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 8 | Sales Tax |
| 2 | Dashboard | 9 | Quarterly Taxes |
| 3 | Schedule C P&L | 10 | Schedule C Lines |
| 4 | Income | 11 | Invoices |
| 5 | COGS & Inventory | 12 | Reconciliation |
| 6 | Expenses | 13 | Monthly Summary |
| 7 | Mileage | 14 | Settings |

## The 12 printable PDF pages

Schedule C P&L · Tax Worksheet · Income Log · Expense Log · COGS & Inventory · Mileage
Log · Sales Tax Log · Invoice Tracker · Receipt Tracker · Reconciliation · Monthly
Summary · Tax-Time Checklist.

---

## Signature automation — profit, then the tax it creates

Everything connects. Revenue minus what it cost to make and run gives your real net
profit, and that profit produces an exact quarterly payment:

```
Gross profit   = revenue − COGS
Net profit     = gross profit − expenses        (Schedule C Line 31)
SE tax         = net profit × 92.35% × 15.3%
Income tax     = (net profit − ½ SE tax) × your rate
Quarterly      = (SE tax + income tax) ÷ 4
```

### The 12 dashboard KPIs
Gross Revenue · COGS · Gross Profit · Expenses · Net Profit · Net Margin · SE Tax ·
Income Tax · Total Tax · Quarterly · Mileage Deduction · Books Score.
The **Books Score** blends profitable, healthy-margin, tax-set-aside,
expenses-categorized, books-reconciled and receipts-attached into one 0–100% number.

**Verified sample business** (Quill & Ledger, owner Morgan): gross revenue **$96,000** ·
COGS **$28,800** · gross profit **$67,200** · expenses **$19,200** · net profit
**$48,000** · net margin **50%** · SE tax **$6,782** · income tax **$5,353** · total tax
**$12,135** (25.3% effective) · quarterly **$3,034** · mileage deduction **$2,800** ·
**Books Score 90%**.

---

## Premium bookkeeping design

- A **Schedule C P&L** engine with COGS and a real Line 31 net profit
- A **quarterly tax** calculator (SE tax + income tax ÷ 4)
- **Mileage** at the IRS rate — the deduction most businesses forget
- **Sales tax** collected vs remitted, and **invoices** outstanding
- Expense categories **mapped to real Schedule C lines** for your CPA
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A bookkeeping & organizing tool, not tax, legal or accounting advice.** Confirm
> every figure with your own tax professional.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Bookkeeping_Tax_Command_Center.xlsx
python3 build_pdf.py                         # -> ../Bookkeeping_Tax_Printables.pdf (+ page PNGs)
python3 build_marketing.py                  # -> ../marketing/01..06.png
python3 build_marketing_detail.py           # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

# Freelancer Cashflow & Tax Command Center™ — The Self-Employed Operating System

> Not an invoice template — a **complete get-paid, set-aside-tax, keep-it system**.
> One premium **Google Sheets + printable PDF** command center for freelancers: a
> cashflow-and-tax engine (income − expenses − auto tax set-aside → your real
> take-home), invoices, clients, time & rates, business expenses, a tax vault, mileage
> & home office, a pipeline, savings & runway, subscriptions and a monthly summary —
> everything cross-linked and live.

| | |
| - | - |
| **Product** | Freelancer Cashflow & Tax Command Center™ |
| **Target** | Freelancers & the self-employed · designers, writers & developers · consultants & coaches · virtual assistants & marketers · 1099 contractors & side-hustlers · anyone who invoices clients |
| **Angle** | Get paid, set aside tax, and know your real hourly rate. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $19 single file · **$29 bundle** (Sheets + PDF) · $49 with a commercial-use file license — all the *same digital files*, never a service |

---

## Contents

```
products/freelancer-command-center/
├── README.md
├── Freelancer_Command_Center.xlsx     ← Google Sheets / Excel master (14 tabs)
├── Freelancer_Printables.pdf          ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 8 | Tax Vault |
| 2 | Dashboard | 9 | Mileage & Home Office |
| 3 | Cashflow & Tax | 10 | Pipeline |
| 4 | Invoices | 11 | Savings & Runway |
| 5 | Clients | 12 | Subscriptions |
| 6 | Time & Rates | 13 | Monthly Summary |
| 7 | Business Expenses | 14 | Settings |

## The 12 printable PDF pages

Invoice · Cashflow & Tax · Client List · Rate Card · Business Expenses · Tax Vault ·
Mileage & Office · Pipeline · Savings & Runway · Subscriptions · Monthly Summary ·
Weekly Checklist.

---

## Signature automation — set the tax aside before you spend it

Everything connects. Your paid invoices are your income, your expenses lower your net,
and the engine sets aside tax automatically so the take-home is truly yours:

```
Income        = Σ paid invoices
Net income    = Income − business expenses
Tax set-aside = Net income × tax rate
Take-home     = Net income − tax set-aside
Effective rate= Net income ÷ billable hours
```

### The 12 dashboard KPIs
Monthly Income · Expenses · Net Income · Tax Set-Aside · Take-Home · Effective Rate ·
Billable Hours · Utilization · Invoices · Outstanding · Year Pace · Freelance Score.
The **Freelance Score** blends income-vs-goal, margin, effective-rate, take-home,
tax-set-aside and runway into one 0–100% number.

**Verified sample freelancer** (Studio Fern, owner Sasha): income **$8,000** ·
expenses **$1,200** · net **$6,800** · tax set-aside **$2,040** · take-home **$4,760**
· effective rate **$68.00**/hr · **100** billable hours · utilization **63%** · **6**
invoices · outstanding **$2,400** · year pace **$96,000** · **Freelance Score 90%**.

---

## Premium freelance design

- A **cashflow-and-tax** engine with automatic tax set-aside & take-home
- **Invoices** (paid vs outstanding) and a **client** roster
- **Time & rates** with your true effective hourly rate
- A **tax vault** with quarterly estimates and **mileage / home-office** deductions
- A **pipeline**, a **savings & runway** tracker and a **subscriptions** audit
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A business tool, not financial, legal or tax advice.** The tax set-aside is an
> estimate — confirm your rate and payments with your own tax advisor.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Freelancer_Command_Center.xlsx
python3 build_pdf.py                         # -> ../Freelancer_Printables.pdf (+ page PNGs)
python3 build_marketing.py                  # -> ../marketing/01..06.png
python3 build_marketing_detail.py           # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

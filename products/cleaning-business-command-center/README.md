# Cleaning & Service Business Command Center™ — The Service-Business Operating System

> Not a price list — a **complete price-it, book-it, grow-the-recurring system**.
> One premium **Google Sheets + printable PDF** command center for cleaners and
> home-service pros: a job-P&L engine (price − supplies − labor − travel → job profit
> & profit per hour), services & pricing, a client roster that becomes recurring
> revenue, a schedule, leads & quotes, supplies, team & labor, mileage, expenses,
> reviews and a monthly summary — everything cross-linked and live.

| | |
| - | - |
| **Product** | Cleaning & Service Business Command Center™ |
| **Target** | House & office cleaners · maid & janitorial services · lawn care & landscapers · window & pressure washers · handymen & home-service pros · any solo or small service crew |
| **Angle** | Price every job, keep the calendar full, and grow recurring revenue. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $19 single file · **$29 bundle** (Sheets + PDF) · $49 with a commercial-use file license — all the *same digital files*, never a service |

---

## Contents

```
products/cleaning-business-command-center/
├── README.md
├── Cleaning_Business_Command_Center.xlsx  ← Google Sheets / Excel master (14 tabs)
├── Cleaning_Business_Printables.pdf       ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 8 | Supplies |
| 2 | Dashboard | 9 | Team & Labor |
| 3 | Job P&L | 10 | Mileage & Vehicle |
| 4 | Services & Pricing | 11 | Expenses |
| 5 | Clients | 12 | Reviews & Referrals |
| 6 | Schedule | 13 | Monthly Summary |
| 7 | Leads & Quotes | 14 | Settings |

## The 12 printable PDF pages

Job Quote & P&L · Service Price List · Client List · Weekly Schedule · Leads & Quotes
· Supply List · Team & Labor · Mileage Log · Expense Log · Reviews & Referrals ·
Monthly Summary · Job Checklist.

---

## Signature automation — did this job pay, and what per hour?

Everything connects. Your job price minus every cost gives the job profit, your hours
turn it into profit per hour, and your recurring clients become monthly revenue:

```
Job costs      = Supplies + labor + travel
Job profit     = Job price − job costs
Profit / hour  = Job profit ÷ hours on the job
Recurring (MRR)= Σ (client price × jobs per month)
Monthly revenue= MRR + one-time jobs
```

### The 12 dashboard KPIs
Job Price · Job Costs · Job Profit · Job Margin · Profit/Hour · Jobs/Month · Monthly
Revenue · Recurring (MRR) · Active Clients · Monthly Profit · Avg Job · Service Score.
The **Service Score** blends job-margin, profit-per-hour, services-priced, profitable,
recurring-revenue and client-growth into one 0–100% number.

**Verified sample business** (Bright & Tidy Co., owner Ava): job price **$180** · job
costs **$80** · job profit **$100** · job margin **56%** · profit/hour **$40.00** ·
**31** jobs/month · monthly revenue **$6,000** · recurring **$4,500** · **12** clients
· monthly profit **$2,900** · avg job **$194** · **Service Score 90%**.

---

## Premium service-business design

- A **job-P&L** engine (job profit and profit per hour on any service)
- **Services & pricing** with the rate per hour on every service
- A **client roster** that turns recurring clients into **MRR**
- A **schedule**, **leads & quotes**, **supplies**, **team** and **mileage**
- **Expenses**, **reviews & referrals** and a **monthly summary**
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A business tool, not financial, legal or tax advice.** Confirm figures with your
> own advisors.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Cleaning_Business_Command_Center.xlsx
python3 build_pdf.py                         # -> ../Cleaning_Business_Printables.pdf (+ page PNGs)
python3 build_marketing.py                  # -> ../marketing/01..06.png
python3 build_marketing_detail.py           # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

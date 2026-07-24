# Subscription & Bills Audit Command Center™ — The Recurring-Spend Operating System

> Not a subscription list — a **complete find-it, cut-it, keep-it-cut system**.
> One premium **Google Sheets + printable PDF** command center for recurring spend: an
> audit engine (every charge → monthly & annualized, with a cancel-savings finder),
> subscriptions, bills, a cancel finder, spend by category, renewals, free trials, price
> hikes, a savings log, a negotiation list and a monthly summary — everything
> cross-linked and live.

| | |
| - | - |
| **Product** | Subscription & Bills Audit Command Center™ |
| **Target** | Anyone with too many subscriptions · budgeters cutting waste · couples sharing bills · frugal & FIRE-minded savers · families trimming expenses |
| **Angle** | See every recurring charge, and cut the ones you forgot. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $19 single file · **$29 bundle** (Sheets + PDF) · $49 with a commercial-use file license — all the *same digital files*, never a service |

---

## Contents

```
products/subscription-audit-command-center/
├── README.md
├── Subscription_Audit_Command_Center.xlsx   ← Google Sheets / Excel master (14 tabs)
├── Subscription_Audit_Printables.pdf         ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 8 | Renewals |
| 2 | Dashboard | 9 | Free Trials |
| 3 | Subscription Audit | 10 | Price Hikes |
| 4 | Subscriptions | 11 | Savings |
| 5 | Bills | 12 | Negotiation |
| 6 | Cancel Finder | 13 | Monthly Summary |
| 7 | Categories | 14 | Settings |

## The 12 printable PDF pages

Subscription Audit · Cancel List · Bills List · Renewal Calendar · Free Trial Tracker ·
Price Hike Log · Spend by Category · Bill Negotiation · Savings Log · Monthly Summary ·
Audit Worksheet · Audit Checklist.

---

## Signature automation — the cancel-savings finder

Everything connects. Your subscriptions total your recurring spend, the ones you flag
'Cancel' total your savings, and everything annualizes:

```
Monthly subs   = Σ subscription monthly costs
Annual subs    = monthly subs × 12
Cancel savings = Σ (subscriptions flagged "Cancel")   → × 12 for annual
Keep monthly   = monthly subs − cancel savings
```

### The 12 dashboard KPIs
Monthly Subs · Annual Subs · Subscriptions · Avg/Sub · Flagged · Monthly Savings ·
Annual Savings · Monthly Bills · Keep Monthly · Hikes Add/Yr · Trials Ending · Audit
Score. The **Audit Score** blends everything-reviewed, cutting-waste, all-categorized,
under-budget, trials-tracked and billed-annually into one 0–100% number.

**Verified sample audit** (Clearing House, Devon): monthly subs **$216.77** · annual
subs **$2,601** · subscriptions **14** · avg/sub **$15.48** · flagged **5** · monthly
savings **$105.92** · annual savings **$1,271** · monthly bills **$468** · keep monthly
**$110.85** · hikes add **$132**/yr · trials ending **3** · **Audit Score 90%**.

---

## Premium audit design

- A **cancel-savings finder** (flag it, see the annual savings instantly)
- Every **subscription** and **bill**, monthly and annualized
- A **cancel finder**, **spend by category** and a **renewal calendar**
- **Free-trial** reminders, **price-hike** tracking and a **negotiation** list
- A **savings log** and a monthly recurring-spend trend that goes down
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A budgeting & organizing tool, not financial advice.** Confirm figures with your
> own statements.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Subscription_Audit_Command_Center.xlsx
python3 build_pdf.py                         # -> ../Subscription_Audit_Printables.pdf (+ page PNGs)
python3 build_marketing.py                  # -> ../marketing/01..06.png
python3 build_marketing_detail.py           # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

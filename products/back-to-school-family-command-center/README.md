# Back-to-School Command Center™ — The Large-Family Back-to-School System

> Not a checklist — a **complete large-family back-to-school operating system**.
> One premium **Google Sheets + printable PDF** command center for every child,
> contact, form, fee, supply, uniform, lunch, deadline and dollar. Built by a
> mom of six, for families of **1 to 8 kids**.

| | |
| - | - |
| **Product** | Back-to-School Command Center™ |
| **Target** | Busy parents of 1–8 kids · large & blended families · multiple schools & schedules · grandparents & caregivers · foster / kinship families · ADHD-household organizers |
| **Angle** | Every child, form, fee & deadline in one calm system — from a mom who actually runs a full house. |
| **Formats** | Google Sheets (17-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $12 single · **$18 bundle** (Sheets + PDF) · $24 with lunch & meal-plan add-on · $59 shop / PLR-style license |

---

## Contents

```
products/back-to-school-family-command-center/
├── README.md
├── Back_to_School_Command_Center.xlsx   ← Google Sheets / Excel master (17 tabs + Settings)
├── Back_to_School_Printables.pdf        ← 12-page print-ready pack (US Letter)
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

## The 17-tab system

| # | Tab | # | Tab |
| - | --- | - | --- |
| 1 | Start Here | 10 | Fees & Payments |
| 2 | Family Dashboard | 11 | Extracurriculars |
| 3 | Child Profiles | 12 | Lunch & Grocery Planner |
| 4 | School & Teacher Contacts | 13 | Homework & Reading Log |
| 5 | 2026–2027 Calendar | 14 | Parent-Teacher Comms Log |
| 6 | Events & Deadlines | 15 | Absence & Late Tracker |
| 7 | Supply Shopping Tracker | 16 | Grades & Report Cards |
| 8 | Clothing & Uniform Inventory | 17 | Important Documents |
| 9 | Budget vs Actual | | *(+ Settings)* |

## The 12 printable PDF pages

Child Information Sheet · School Contact Page · First-Day Prep Checklist ·
Backpack Checklist · Weekly Family Schedule · School Supply Checklist ·
Clothing-Size Tracker · Lunchbox Planner · Field-Trip & Payment Log ·
Parent-Teacher Meeting Notes · School-Year Goals · First-Day & Last-Day Memory Pages

---

## Signature automation (12 dashboard KPIs)

| Metric | How it's computed |
| ------ | ----------------- |
| Children | `=Children` |
| Schools | `=Schools` |
| First Day | `=FirstDay` |
| Supplies Bought | `=COUNTIF(SupBought,"Yes")/COUNTA(SupItem)` |
| Budget Spent | `=BudgetSpent` |
| Budget Left | `=BudgetPlanTotal-BudgetSpent` |
| Fees Paid | `=FeePaidTotal/FeeAmtTotal` |
| Forms Done | `=COUNTIF(DocDone,"Yes")/COUNTA(DocName)` |
| Uniforms Ready | `=COUNTIF(ClothReady,"Ready")/COUNTA(ClothItem)` |
| Events (30d) | `=COUNTIFS(EventDate…TODAY()…+30)` |
| To-Do Open | `=COUNTIF(EventStatus,"To Do")+…"In Progress")` |
| Readiness Score | `=AVERAGE(HealthRange)` |

Checking off a supply, paying a fee or marking a form updates the budget and the
**Readiness Score** live; the score blends supplies, clothing & uniforms, fees,
forms, overdue-caught-up and budget-on-track into one 0–100% "how ready are we"
number.

**Verified sample family** (the Riveras — 6 kids across 4 schools):
Children **6** · Schools **4** · First day **Aug 25** · Supplies **85%** bought ·
Budget **$1,850** spent of **$2,500** (**$650** left) · Fees **88%** paid ·
Forms **80%** done · Uniforms **75%** ready · Events **9** in 30 days · To-do
**10** · **Readiness 82%**.

---

## Premium family-software design

- Two-row **gold-divider headers** on every tab; a true Family Dashboard (12
  KPIs + readiness bars, a budget donut, a "due this week" list & a readiness gauge)
- Status color-coding (Done / In Progress / To Do; Ready / Ordered / Need),
  data-bars and conditional flags throughout
- **Print-ready PDF pack** on white with a forest-green header band & gold
  rules — ink-light and fillable
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Back_to_School_Command_Center.xlsx
python3 build_pdf.py                        # -> ../Back_to_School_Printables.pdf (+ page PNGs)
python3 build_marketing.py                 # -> ../marketing/01..06.png
python3 build_marketing_detail.py          # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

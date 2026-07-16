# High-School Transcript Command Center™ — GPA, Credits & a Print-Ready Official Transcript

> Not a blank transcript template — a **complete high-school records system** that
> does the math for you. One premium **Google Sheets + printable PDF** command
> center for course records, auto-calculating GPA (weighted *and* unweighted),
> credits, graduation requirements, test scores, activities, awards, service and
> a clean, official-looking transcript you can send with applications.

| | |
| - | - |
| **Product** | High-School Transcript Command Center™ |
| **Target** | Homeschool & college-bound families · homeschool high-schoolers · transcript-keeping parents · guidance-minded moms · dual-enrollment students · anyone building a college application file |
| **Angle** | Enter grades once — the transcript, GPA and credits do the rest. |
| **Formats** | Google Sheets (17-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $16 single · **$24 bundle** (Sheets + PDF) · $32 with 4-year planning add-on · $79 counselor / commercial license |

---

## Contents

```
products/highschool-transcript-command-center/
├── README.md
├── HS_Transcript_Command_Center.xlsx   ← Google Sheets / Excel master (17 tabs)
├── HS_Transcript_Printables.pdf        ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 10 | Activities |
| 2 | Dashboard | 11 | Awards |
| 3 | Student Profile | 12 | Service |
| 4 | Official Transcript | 13 | Course Descriptions |
| 5 | Course Records | 14 | Reading List |
| 6 | GPA Calculator | 15 | Grad Requirements |
| 7 | Credit Tracker | 16 | Portfolio |
| 8 | 4-Year Plan | 17 | Settings |
| 9 | Test Scores | | |

## The 12 printable PDF pages

Official Transcript · 4-Year Plan · Course Records / Grade Sheet · GPA Worksheet ·
Credit & Graduation Checklist · Test-Score Log · Awards & Honors · Service-Hours
Log · Course Descriptions · Reading List · College Application Tracker ·
Activities Résumé.

---

## Signature automation — GPA that calculates itself

Type a **letter grade** in Course Records and everything downstream computes:

| Column | Formula (per course row *r*) |
| ------ | ---------------------------- |
| Grade points | `=IFERROR(VLOOKUP(F{r},GradeScale,2,FALSE),"")` |
| Quality points (unweighted) | `=IFERROR(G{r}*E{r},"")` |
| Weighted points | `=IFERROR((G{r}+IFERROR(VLOOKUP(D{r},LevelScale,2,FALSE),0))*E{r},"")` |

`LevelScale` adds the honors bump automatically — Regular +0.0, Honors +0.5,
AP +1.0, Dual Credit +1.0 — so a weighted GPA appears **without a single manual
calculation**.

### The 12 dashboard KPIs

| Metric | How it's computed |
| ------ | ----------------- |
| GPA (Weighted) | `=GPAW` (`=WtdTotal/CreditsEarned`) |
| GPA (Unweighted) | `=GPAUW` (`=QualTotal/CreditsEarned`) |
| Credits Earned | `=CreditsEarned` |
| Grad Progress | `=MIN(CreditsEarned/CreditsReq,1)` |
| Courses | `=CourseCount` |
| Honors / AP | `=HonorsAP` |
| Class Of | `=ClassOf` |
| Best SAT | best SAT on file |
| Best ACT | best ACT on file |
| Service Hrs | `=ServiceHours` |
| Activities | `=COUNTA(ActName)` |
| College-Ready | `=AVERAGE(HealthRange)` |

The **College-Ready Score** blends GPA-vs-target, credits, graduation
requirements met, test scores on file, course rigor and service/activities into
one 0–100% "is the file application-ready?" number.

**Verified sample student** (Ella Bennett, Class of 2027): Courses **27** ·
Credits **24.0** of 24 · Unweighted GPA **3.81** · Weighted GPA **4.06** ·
Honors/AP **9** · Best SAT **1380** · Best ACT **30** · Service **120** hrs ·
Activities **7** · Awards **5** · Grad requirements **9 of 10** ·
**College-Ready 98%**.

---

## Premium records-software design

- A true dashboard: 12 KPIs, a GPA-by-year bar chart, a credits-by-subject
  breakdown, a college-application file list and a College-Ready gauge
- **Auto-GPA** by year and cumulative — weighted *and* unweighted — driven by a
  `GradeScale` + `LevelScale` VLOOKUP engine
- A print-ready **Official Transcript** tab (school header, per-year blocks,
  summary box, signature line) that matches the PDF page exactly
- Credit-by-subject tracking against graduation requirements, with color-coded
  status flags
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **Not an accredited transcript service.** Graduation and diploma requirements
> vary by state and by the colleges you apply to — always confirm your own
> jurisdiction's and target schools' rules. This is a planning &
> record-keeping tool.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../HS_Transcript_Command_Center.xlsx
python3 build_pdf.py                        # -> ../HS_Transcript_Printables.pdf (+ page PNGs)
python3 build_marketing.py                 # -> ../marketing/01..06.png
python3 build_marketing_detail.py          # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

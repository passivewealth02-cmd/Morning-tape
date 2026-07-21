# College Application Command Center™ — The Complete Admissions System

> Not a checklist — a **complete apply-smart, compare-offers admissions system**.
> One premium **Google Sheets + printable PDF** command center for the whole
> senior-year application season: a balanced college list (reach / match / safety)
> with per-school progress, an essay & supplement tracker, recommendations, test
> scores, an activities résumé, scholarships, a financial-aid / net-price
> comparison, visits & interviews, a decisions tracker, a master to-do list and a
> deadlines calendar.

| | |
| - | - |
| **Product** | College Application Command Center™ |
| **Target** | High-school seniors & juniors · homeschool & transfer applicants · parents helping with admissions · school counselors · first-gen & scholarship-focused students · anyone applying to multiple colleges |
| **Angle** | Apply smart, hit every deadline, compare offers — the whole season, organized. |
| **Formats** | Google Sheets (15-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $16 single · **$24 bundle** (Sheets + PDF) · $32 with the transcript add-on · $79 counselor / commercial license |

---

## Contents

```
products/college-command-center/
├── README.md
├── College_Command_Center.xlsx   ← Google Sheets / Excel master (15 tabs)
├── College_Printables.pdf        ← 12-page print-ready pack (US Letter)
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

## The 15-tab system

| # | Tab | # | Tab |
| - | --- | - | --- |
| 1 | Start Here | 9 | Scholarships |
| 2 | Dashboard | 10 | Net Price |
| 3 | Applicant Profile | 11 | Visits |
| 4 | College List | 12 | Decisions |
| 5 | Essays | 13 | To-Do |
| 6 | Recommendations | 14 | Deadlines |
| 7 | Test Scores | 15 | Settings |
| 8 | Activities | | |

## The 12 printable PDF pages

Applicant Profile · College List · Application Tracker · Essay & Supplement
Tracker · Recommendation Tracker · Test Scores · Activities & Awards Résumé ·
Scholarship Log · Net-Price Comparison · Visits & Interviews · Decisions &
Compare · Master To-Do & Deadlines.

---

## Signature automation — a college list that tracks itself

Check off each school's **essays, recs, form & submission** and per-school
progress computes: `=COUNTIF(E:H,"Yes")/4`.

### The 12 dashboard KPIs
Colleges · Apps Submitted · Avg Progress · Essays Done · Recs Secured ·
Scholarships · Aid Awarded · Next Deadline · Acceptances · Best Net Price · Tasks
Done · Ready Score. The **Ready Score** blends application progress, apps
submitted, essays final, recs secured, scholarships applied and tasks done into
one 0–100% number, and a **Reach / Match / Safety** donut keeps the list
balanced.

**Verified sample applicant** (Ella Bennett, Class of 2027 — the same student as
the transcript product): Colleges **8** (3 reach / 3 match / 2 safety) · Apps
submitted **4** · Avg progress **69%** · Essays final **50%** · Recs **3** of 4 ·
Scholarships applied **6** · Aid awarded **$12,000** · Acceptances **2** · Best
net price **$2,000**/yr · Tasks **60%** · **Ready Score 63%**.

---

## Premium admissions-software design

- A **balanced college list** with reach/match/safety flags & auto progress bars
- **Essay & recommendation** trackers so nothing is late; a **scholarship** log
- A **net-price comparison** that highlights the cheapest true cost after aid
- Visits, interviews, a **decisions** tracker and an in-order **deadlines** view
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **An organizing tool, not admissions advice.** Confirm every deadline, fee,
> policy and requirement directly with each college — they change, and each
> school is the final word.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../College_Command_Center.xlsx
python3 build_pdf.py                        # -> ../College_Printables.pdf (+ page PNGs)
python3 build_marketing.py                 # -> ../marketing/01..06.png
python3 build_marketing_detail.py          # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

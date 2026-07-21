# IEP & Special-Needs Command Center™ — The Complete Advocacy & Progress System

> Not a folder of paperwork — a **complete organizing, tracking & advocacy
> system**. One calm **Google Sheets + printable PDF** command center for a
> child's IEP/504: goals with real progress monitoring, services & minutes,
> accommodations, a therapy log, a gentle behavior tracker, meeting &
> communication logs, health/medication notes, a strengths profile, a records
> checklist and a wins log. Built by-a-parent, for parents.

| | |
| - | - |
| **Product** | IEP & Special-Needs Command Center™ |
| **Target** | Parents & caregivers of kids with an IEP or 504 · newly-diagnosed & first-IEP families · foster & kinship caregivers · advocates & parent mentors · families managing therapies & services |
| **Angle** | Walk into every meeting organized — advocate with confidence. |
| **Formats** | Google Sheets (16-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $15 single · **$24 bundle** (Sheets + PDF) · $32 with the sibling / multi-child add-on · $79 advocate / commercial license |

---

## Contents

```
products/iep-command-center/
├── README.md
├── IEP_Command_Center.xlsx     ← Google Sheets / Excel master (16 tabs)
├── IEP_Printables.pdf          ← 12-page printable advocacy binder (US Letter)
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

## The 16-tab system

| # | Tab | # | Tab |
| - | --- | - | --- |
| 1 | Start Here | 9 | Behavior |
| 2 | Dashboard | 10 | Meetings |
| 3 | Student Profile | 11 | Communication |
| 4 | IEP Goals | 12 | Health & Meds |
| 5 | Progress Monitoring | 13 | Strengths |
| 6 | Services | 14 | Records |
| 7 | Accommodations | 15 | Wins |
| 8 | Therapy Log | 16 | Settings |

## The 12 printable PDF pages

Student Profile & Team · IEP Goals & Progress · Progress Monitoring · Services &
Minutes · Accommodations Checklist · Therapy / Session Log · Behavior Tracker
(ABC) · Meeting Notes / Prep · Communication Log · Strengths & Interests ·
Records & Documents · Wins & Milestones.

---

## Signature automation — goals that trend themselves

Enter each goal's **baseline, target & current** level and progress calculates:

```
Progress = MIN( MAX( (current − baseline) / (target − baseline), 0 ), 1 )
```

### The 12 dashboard KPIs
IEP Goals · Avg Progress · Goals On Pace · Services · Service Min/Wk ·
Accommodations · Therapy Logged · Data Points · Meetings · Wins · Supports in
Place · Progress Score. The **Progress Score** blends goal progress, on-pace
goals, accommodations in place, services delivered vs scheduled, data collection
and records-ready into one 0–100% number, and a **Progress Toward Each Goal**
bar chart shows every goal at a glance.

**Verified sample child** (fictional — "Sam", grade 3): IEP goals **5** · Avg
progress **64%** · Goals on pace **4** of 5 · Services **4** · Service minutes
**420**/wk (395 delivered) · Accommodations **9** of 10 · Therapy logged **8** ·
Data points **24** · Meetings **3** · Wins **6** · Supports in place **90%** ·
**Progress Score 80%**.

---

## Premium, parent-first design

- IEP goals with real **progress monitoring** — see exactly how each goal trends
- Services & minutes: what's **owed vs delivered**; accommodations: **is it in place?**
- Therapy, behavior (gentle ABC), meeting & dated **communication** logs
- A **strengths-based** profile to open every meeting, and a **Wins** log for the
  progress that gets lost between reviews
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **Not medical, legal, psychological or educational advice.** This is a personal
> organizing & advocacy tool; it does not create or replace an IEP/504. Always
> work with your child's IEP team and qualified professionals, and follow your
> district's process. Keep the file private — it may contain sensitive information.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../IEP_Command_Center.xlsx
python3 build_pdf.py                        # -> ../IEP_Printables.pdf (+ page PNGs)
python3 build_marketing.py                 # -> ../marketing/01..06.png
python3 build_marketing_detail.py          # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.

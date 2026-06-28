# Dividend Wealth Builder System (DWBS)

> Track. Optimize. Forecast your way to financial independence.

| | |
| - | - |
| **Product** | Dividend Wealth Builder System |
| **Target** | Long-term investors, dividend portfolio builders, FIRE planners, retirement planners, passive-income seekers |
| **Angle** | "Build $1,000 → $10,000+ monthly passive dividend income with a structured portfolio tracking system." |
| **Formats** | Excel `.xlsx` + Google Sheets edition |
| **Pricing** | $19 (single) · **$29 bundle** · $79 with onboarding |

---

## Contents

```
products/dividend-wealth-builder/
├── README.md                       ← this file
├── Dividend_Wealth_Builder.xlsx    ← the Excel product (ship this)
├── GOOGLE_SHEETS.md                ← Google Sheets formulas & setup
├── BUILD_INSTRUCTIONS.md           ← step-by-step build & ship guide
└── build/
    └── build_xlsx.py               ← deterministic .xlsx generator
```

---

## Workbook architecture

| Sheet | Role | Tab color |
| ----- | ---- | --------- |
| **Dashboard** | KPI row · sector pie · monthly bar · wealth curve · yield-vs-risk scatter · filters | Primary `#1B4F48` |
| **Holdings** | 60-row database, validated, conditionally formatted | Accent `#937356` |
| **Calculations** | Centralized formulas: aggregates, sector breakdown, monthly distribution | Highlight `#75E6C1` |
| **Settings** | User inputs, named ranges, dropdown lists, frequency lookup | Surface `#E5D3BA` |
| **Projections** | FIRE metrics + 30-year compounded projection + 3 charts | Primary `#1B4F48` |

---

## Formula engine highlights

- `SUMPRODUCT` / `SUMIFS` portfolio aggregation
- `VLOOKUP` into a `FreqTable` for payments-per-year resolution
- 30-year recursive projection with optional reinvestment toggle
- `LOG`-based years-to-target solver for $1k, $5k, and custom targets
- Six dynamic charts, all hooked to named ranges (data extends → charts grow)
- Sector ColorScale, Risk DataBar, loss-row FormulaRule conditional formats
- Three data-validation lists driven from Settings (Sector / Frequency / Risk)

---

## Brand system applied

| Token | Hex | Where it shows |
| ----- | --- | -------------- |
| Primary | `#1B4F48` | Header bars, table headers, KPI values |
| Accent | `#937356` | KPI labels, warning indicators |
| Surface | `#E5D3BA` | Input cells, row banding |
| Highlight | `#75E6C1` | Yield top-end color scale, positive KPIs |
| Danger | `#C94C4C` | Risk bar, loss-row highlight |
| Background | `#FFFFFF` | All cards, body |
| Text | `#333333` | Body copy |

---

## Build & ship

```bash
cd build && python3 build_xlsx.py
```

Output: `../Dividend_Wealth_Builder.xlsx`. See `BUILD_INSTRUCTIONS.md`
for the Google Sheets path and Etsy delivery package.

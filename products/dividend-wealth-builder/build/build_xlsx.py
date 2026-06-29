"""Build the Dividend Wealth Builder System (DWBS) Excel workbook — PREMIUM.

6 sheets · luxury dividend operating system.
  Dashboard · Holdings · Income Calendar · Calculations · Projections · Settings

Run: python3 build_xlsx.py
Outputs: ../Dividend_Wealth_Builder.xlsx
"""
from __future__ import annotations

import os

from openpyxl import Workbook
from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference, ScatterChart, Series
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, NamedStyle, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

# ---------------------------------------------------------------------------
# Brand tokens (+ luxury extensions)
# ---------------------------------------------------------------------------
PRIMARY = "1B4F48"
PRIMARY_DK = "133A35"
ACCENT = "937356"
GOLD_LT = "C9A86A"
SURFACE = "E5D3BA"
HIGHLIGHT = "75E6C1"
MINT_BG = "E3F8EF"
BG = "FFFFFF"
WHITE = "FFFFFF"
TEXT = "333333"
DANGER = "C94C4C"
RED_BG = "FBE6E6"
WARN_BG = "FBF0E2"
MUTED_ROW = "F4ECDE"
BORDER = "D6D2C8"
SOFT_BG = "FAF7F1"
IVORY = "FBF8F2"

SECTORS = ["Tech", "Finance", "Energy", "REITs", "Consumer", "Healthcare"]
RISK_LEVELS = ["Low", "Medium", "High"]
FREQUENCIES = ["Monthly", "Quarterly", "Semi-Annual", "Annual"]
FREQ_PAYMENTS = {"Monthly": 12, "Quarterly": 4, "Semi-Annual": 2, "Annual": 1}

THIN = Side(style="thin", color=BORDER)
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
GOLD = Side(style="medium", color=GOLD_LT)

SAMPLE_HOLDINGS = [
    # ticker, company, sector, shares, avg_cost, price, div/sh, frequency, risk
    ("SCHD", "Schwab US Dividend ETF", "Finance", 220, 72.30, 78.40, 2.66, "Quarterly", 2),
    ("O", "Realty Income", "REITs", 200, 62.40, 58.10, 3.16, "Monthly", 2),
    ("JNJ", "Johnson & Johnson", "Healthcare", 60, 162.45, 158.20, 4.76, "Quarterly", 1),
    ("MSFT", "Microsoft Corp.", "Tech", 40, 250.10, 415.20, 3.00, "Quarterly", 2),
    ("KO", "Coca-Cola Co.", "Consumer", 120, 56.10, 63.40, 1.84, "Quarterly", 1),
    ("PG", "Procter & Gamble", "Consumer", 45, 142.20, 161.80, 4.03, "Quarterly", 1),
    ("JPM", "JPMorgan Chase", "Finance", 35, 132.50, 198.10, 4.60, "Quarterly", 3),
    ("VYM", "Vanguard High Div ETF", "Finance", 100, 105.40, 124.80, 3.45, "Quarterly", 2),
    ("XOM", "Exxon Mobil", "Energy", 80, 95.20, 116.30, 3.80, "Quarterly", 3),
    ("CVX", "Chevron Corp.", "Energy", 30, 158.10, 162.40, 6.52, "Quarterly", 3),
    ("AAPL", "Apple Inc.", "Tech", 50, 145.30, 192.40, 1.00, "Quarterly", 2),
    ("MAIN", "Main Street Capital", "Finance", 90, 41.10, 49.80, 2.94, "Monthly", 3),
    ("STAG", "Stag Industrial", "REITs", 180, 35.20, 36.85, 1.48, "Monthly", 3),
    ("MO", "Altria Group", "Consumer", 110, 45.60, 49.20, 3.92, "Quarterly", 3),
    ("ABBV", "AbbVie Inc.", "Healthcare", 40, 138.40, 176.20, 6.20, "Quarterly", 2),
]
MAX_ROWS = 40
H_START = 5
H_END = H_START + MAX_ROWS - 1


# ===========================================================================
# Styles & helpers
# ===========================================================================
def register_styles(wb):
    def f(size, bold=False, color=TEXT, italic=False):
        return Font(name="Calibri", size=size, bold=bold, color=color, italic=italic)

    styles = {
        "title": NamedStyle(name="title", font=f(24, True, "FFFFFF"),
                            fill=PatternFill("solid", fgColor=PRIMARY),
                            alignment=Alignment(horizontal="left", vertical="center", indent=2)),
        "subtitle": NamedStyle(name="subtitle", font=f(11, False, "E5D3BA", italic=True),
                               fill=PatternFill("solid", fgColor=PRIMARY),
                               alignment=Alignment(horizontal="left", vertical="center", indent=2)),
        "section": NamedStyle(name="section", font=f(12, True, PRIMARY),
                              alignment=Alignment(horizontal="left", vertical="center")),
        "section_gold": NamedStyle(name="section_gold", font=f(12, True, ACCENT),
                                   alignment=Alignment(horizontal="left", vertical="center")),
        "th": NamedStyle(name="th", font=f(11, True, "FFFFFF"),
                         fill=PatternFill("solid", fgColor=PRIMARY),
                         alignment=Alignment(horizontal="center", vertical="center", wrap_text=True),
                         border=BOX),
        "td": NamedStyle(name="td", font=f(11, False, TEXT),
                         alignment=Alignment(horizontal="center", vertical="center"), border=BOX),
        "td_left": NamedStyle(name="td_left", font=f(11, False, TEXT),
                              alignment=Alignment(horizontal="left", vertical="center", indent=1),
                              border=BOX),
        "input": NamedStyle(name="input", font=f(11, True, PRIMARY),
                            fill=PatternFill("solid", fgColor=SURFACE),
                            alignment=Alignment(horizontal="center", vertical="center"), border=BOX),
        "field_label": NamedStyle(name="field_label", font=f(10, True, ACCENT),
                                  alignment=Alignment(horizontal="left", vertical="center", indent=1),
                                  border=BOX, fill=PatternFill("solid", fgColor=SOFT_BG)),
        "field_value": NamedStyle(name="field_value", font=f(11, True, PRIMARY),
                                  alignment=Alignment(horizontal="left", vertical="center", indent=1),
                                  border=BOX),
    }
    for s in styles.values():
        if s.name not in wb.named_styles:
            wb.add_named_style(s)


def fill(c):
    return PatternFill("solid", fgColor=c)


def merge_set(ws, rng, value, style):
    ws.merge_cells(rng)
    cell = ws[rng.split(":")[0]]
    cell.value = value
    cell.style = style
    return cell


def luxe_header(ws, last_col, title, subtitle):
    ws.row_dimensions[1].height = 46
    ws.row_dimensions[2].height = 22
    ws.row_dimensions[3].height = 6
    merge_set(ws, f"A1:{last_col}1", "  " + title, "title")
    merge_set(ws, f"A2:{last_col}2", "  " + subtitle, "subtitle")
    from openpyxl.utils import column_index_from_string
    for c in range(1, column_index_from_string(last_col) + 1):
        ws.cell(row=3, column=c).fill = fill(GOLD_LT)


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def add_dv(ws, rng, list_name):
    dv = DataValidation(type="list", formula1=f"={list_name}", allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(rng)


def kpi_card(ws, row, col, span, label, formula, kind="num"):
    """Premium KPI card: gold-topped white card, label over big value."""
    L, R = get_column_letter(col), get_column_letter(col + span - 1)
    ws.merge_cells(f"{L}{row}:{R}{row}")
    ws.merge_cells(f"{L}{row+1}:{R}{row+1}")
    lc = ws[f"{L}{row}"]
    lc.value = label
    lc.font = Font(size=10, bold=True, color=ACCENT)
    lc.alignment = Alignment(horizontal="center", vertical="center")
    vc = ws[f"{L}{row+1}"]
    vc.value = formula
    vc.font = Font(size=24, bold=True, color=PRIMARY)
    vc.alignment = Alignment(horizontal="center", vertical="center")
    vc.number_format = {"money": '"$"#,##0', "money2": '"$"#,##0.00',
                        "pct": "0.0%", "num": "General"}[kind]
    for rr in (row, row + 1):
        for cc in range(col, col + span):
            c = ws.cell(row=rr, column=cc)
            c.fill = fill(WHITE)
            c.border = Border(left=THIN, right=THIN,
                              top=GOLD if rr == row else THIN, bottom=THIN)
    ws.row_dimensions[row].height = 24
    ws.row_dimensions[row + 1].height = 52


# ===========================================================================
# Settings
# ===========================================================================
def build_settings(wb):
    ws = wb.create_sheet("Settings")
    ws.sheet_view.showGridLines = False
    set_widths(ws, [2, 32, 20, 4, 22, 18, 18, 18])
    luxe_header(ws, "H", "⚙  SETTINGS", "Set your goals once — the whole system updates.")

    merge_set(ws, "B5:C5", "INVESTOR CONTROLS", "section")
    controls = [
        ("Monthly Contribution", 1000, '"$"#,##0', "MonthlyContribution"),
        ("Target Monthly Dividend Income", 4000, '"$"#,##0', "TargetMonthlyIncome"),
        ("Expected Annual Growth Rate", 0.08, "0.00%", "AnnualGrowthRate"),
        ("Dividend Growth Rate", 0.07, "0.00%", "DividendGrowthRate"),
        ("Reinvest Dividends (1=Yes, 0=No)", 1, "0", "ReinvestToggle"),
        ("Dividend Tax Rate", 0.15, "0.00%", "TaxRate"),
        ("Years to Project", 30, "0", "YearsToProject"),
    ]
    for i, (lab, val, fmt, name) in enumerate(controls):
        r = 6 + i
        ws.cell(row=r, column=2, value=lab).style = "field_label"
        c = ws.cell(row=r, column=3, value=val)
        c.style = "input"
        c.number_format = fmt
        wb.defined_names[name] = DefinedName(name, attr_text=f"Settings!$C${r}")

    merge_set(ws, "E5:H5", "DROPDOWN LISTS", "section_gold")
    for col, header, data, name in [
        ("E", "Sectors", SECTORS, "SectorList"),
        ("F", "Risk Levels", RISK_LEVELS, "RiskList"),
        ("G", "Frequencies", FREQUENCIES, "FrequencyList"),
    ]:
        ci = ord(col) - 64
        ws.cell(row=6, column=ci, value=header).style = "th"
        for ri, v in enumerate(data):
            ws.cell(row=7 + ri, column=ci, value=v).style = "td_left"
        wb.defined_names[name] = DefinedName(
            name, attr_text=f"Settings!${col}$7:${col}${6 + len(data)}")

    # Frequency → payments-per-year lookup
    merge_set(ws, "E15:F15", "FREQUENCY LOOKUP", "section_gold")
    ws.cell(row=16, column=5, value="Frequency").style = "th"
    ws.cell(row=16, column=6, value="Pmts/Yr").style = "th"
    for i, fr in enumerate(FREQUENCIES):
        ws.cell(row=17 + i, column=5, value=fr).style = "td_left"
        ws.cell(row=17 + i, column=6, value=FREQ_PAYMENTS[fr]).style = "td"
    wb.defined_names["FreqTable"] = DefinedName(
        "FreqTable", attr_text=f"Settings!$E$17:$F${16 + len(FREQUENCIES)}")


# ===========================================================================
# Holdings
# ===========================================================================
def build_holdings(wb):
    ws = wb.create_sheet("Holdings")
    ws.sheet_view.showGridLines = False
    widths = [10, 26, 13, 9, 12, 12, 14, 14, 13, 10, 12, 13, 11, 10, 10, 9, 8]
    set_widths(ws, widths)
    luxe_header(ws, "Q", "📊  HOLDINGS",
                "Your dividend portfolio — gain/loss, yield, and yield-on-cost auto-calculated.")
    headers = ["Ticker", "Company", "Sector", "Shares", "Avg Cost", "Price",
               "Cost Basis", "Mkt Value", "Gain $", "Gain %", "Div/Share",
               "Frequency", "Annual Inc", "Yield", "YoC", "Weight", "Risk"]
    for i, h in enumerate(headers, 1):
        ws.cell(row=4, column=i, value=h).style = "th"
    ws.row_dimensions[4].height = 30

    for i, h in enumerate(SAMPLE_HOLDINGS):
        r = H_START + i
        tk, co, sec, sh, cost, px, dps, fr, risk = h
        ws.cell(row=r, column=1, value=tk)
        ws.cell(row=r, column=2, value=co)
        ws.cell(row=r, column=3, value=sec)
        ws.cell(row=r, column=4, value=sh)
        ws.cell(row=r, column=5, value=cost)
        ws.cell(row=r, column=6, value=px)
        ws.cell(row=r, column=11, value=dps)
        ws.cell(row=r, column=12, value=fr)
        ws.cell(row=r, column=17, value=risk)

    for r in range(H_START, H_END + 1):
        ws.cell(row=r, column=7, value=f'=IF($D{r}="","",$D{r}*$E{r})')
        ws.cell(row=r, column=8, value=f'=IF($D{r}="","",$D{r}*$F{r})')
        ws.cell(row=r, column=9, value=f'=IF($D{r}="","",$H{r}-$G{r})')
        ws.cell(row=r, column=10, value=f'=IFERROR($I{r}/$G{r},"")')
        ws.cell(row=r, column=13,
                value=(f'=IF(OR($D{r}="",$K{r}="",$L{r}=""),"",'
                       f'$D{r}*$K{r}*VLOOKUP($L{r},FreqTable,2,FALSE))'))
        ws.cell(row=r, column=14,
                value=(f'=IFERROR(IF(OR($F{r}="",$K{r}="",$L{r}=""),"",'
                       f'($K{r}*VLOOKUP($L{r},FreqTable,2,FALSE))/$F{r}),"")'))
        ws.cell(row=r, column=15,
                value=(f'=IFERROR(IF(OR($E{r}="",$K{r}="",$L{r}=""),"",'
                       f'($K{r}*VLOOKUP($L{r},FreqTable,2,FALSE))/$E{r}),"")'))
        ws.cell(row=r, column=16, value=f'=IFERROR($H{r}/TotalMarketValue,"")')

    money = {5, 6, 7, 8, 9, 11, 13}
    pct = {10, 14, 15, 16}
    ints = {4, 17}
    for r in range(H_START, H_END + 1):
        for c in range(1, 18):
            cell = ws.cell(row=r, column=c)
            cell.style = "td_left" if c in (1, 2) else "td"
            cell.fill = fill(MUTED_ROW if (r - H_START) % 2 else BG)
            if c in money:
                cell.number_format = '"$"#,##0.00'
            elif c in pct:
                cell.number_format = "0.00%"
            elif c in ints:
                cell.number_format = "0"
        ws.cell(row=r, column=1).font = Font(bold=True, color=PRIMARY)

    add_dv(ws, f"C{H_START}:C{H_END}", "SectorList")
    add_dv(ws, f"L{H_START}:L{H_END}", "FrequencyList")
    dv_risk = DataValidation(type="whole", operator="between", formula1=1, formula2=5, allow_blank=True)
    ws.add_data_validation(dv_risk)
    dv_risk.add(f"Q{H_START}:Q{H_END}")

    # Conditional formatting
    ws.conditional_formatting.add(
        f"J{H_START}:J{H_END}",
        ColorScaleRule(start_type="num", start_value=-0.3, start_color="FF" + RED_BG,
                       mid_type="num", mid_value=0, mid_color="FFFFF3CD",
                       end_type="num", end_value=0.5, end_color="FF" + HIGHLIGHT))
    ws.conditional_formatting.add(
        f"N{H_START}:N{H_END}",
        ColorScaleRule(start_type="min", start_color="FF" + SURFACE,
                       mid_type="percentile", mid_value=50, mid_color="FFFFE08C",
                       end_type="max", end_color="FF" + HIGHLIGHT))
    ws.conditional_formatting.add(
        f"Q{H_START}:Q{H_END}",
        DataBarRule(start_type="num", start_value=1, end_type="num", end_value=5,
                    color=DANGER, showValue=True))
    ws.conditional_formatting.add(
        f"I{H_START}:I{H_END}",
        CellIsRule(operator="lessThan", formula=["0"],
                   font=Font(color=DANGER, bold=True)))

    # Named ranges
    nm = {
        "TblTicker": "A", "TblCompany": "B", "TblSector": "C", "TblCost": "G",
        "TblMV": "H", "TblGain": "I", "TblIncome": "M", "TblYield": "N",
        "TblRisk": "Q",
    }
    for name, col in nm.items():
        wb.defined_names[name] = DefinedName(
            name, attr_text=f"Holdings!${col}${H_START}:${col}${H_END}")
    ws.freeze_panes = "C5"


# ===========================================================================
# Calculations
# ===========================================================================
def build_calculations(wb):
    ws = wb.create_sheet("Calculations")
    ws.sheet_view.showGridLines = False
    set_widths(ws, [2, 30, 18, 4, 18, 16, 16, 16, 2])
    luxe_header(ws, "H", "🧮  CALCULATIONS ENGINE",
                "Centralized portfolio math. Do not edit — formulas only.")

    merge_set(ws, "B5:C5", "PORTFOLIO TOTALS", "section")
    totals = [
        ("Total Market Value", "=SUM(TblMV)", '"$"#,##0', "TotalMarketValue"),
        ("Total Cost Basis", "=SUM(TblCost)", '"$"#,##0', "TotalCostBasis"),
        ("Total Gain / Loss", "=TotalMarketValue-TotalCostBasis",
         '"$"#,##0;[Red]-"$"#,##0', "TotalGain"),
        ("Total Return %", "=IFERROR(TotalGain/TotalCostBasis,0)", "0.0%", "TotalReturnPct"),
        ("Annual Dividend Income", "=SUM(TblIncome)", '"$"#,##0', "TotalAnnualIncome"),
        ("Monthly Dividend Income", "=TotalAnnualIncome/12", '"$"#,##0.00', "TotalMonthlyIncome"),
        ("Portfolio Yield", "=IFERROR(TotalAnnualIncome/TotalMarketValue,0)", "0.00%", "PortfolioYield"),
        ("Yield on Cost", "=IFERROR(TotalAnnualIncome/TotalCostBasis,0)", "0.00%", "PortfolioYoC"),
        ("After-Tax Monthly", "=TotalMonthlyIncome*(1-TaxRate)", '"$"#,##0.00', "AfterTaxMonthly"),
        ("Holdings Count", "=COUNTA(TblTicker)", "0", "HoldingsCount"),
        ("Progress to Income Goal", "=IFERROR(TotalMonthlyIncome/TargetMonthlyIncome,0)",
         "0.0%", "ProgressToGoal"),
    ]
    for i, (lab, fml, fmt, name) in enumerate(totals):
        r = 6 + i
        ws.cell(row=r, column=2, value=lab).style = "field_label"
        c = ws.cell(row=r, column=3, value=fml)
        c.style = "field_value"
        c.number_format = fmt
        wb.defined_names[name] = DefinedName(name, attr_text=f"Calculations!$C${r}")

    # Sector breakdown
    merge_set(ws, "E5:H5", "SECTOR BREAKDOWN", "section_gold")
    for j, h in enumerate(["Sector", "Value", "Annual Inc", "Weight"]):
        ws.cell(row=6, column=5 + j, value=h).style = "th"
    for i, sec in enumerate(SECTORS):
        r = 7 + i
        ws.cell(row=r, column=5, value=sec).style = "td_left"
        v = ws.cell(row=r, column=6, value=f'=SUMIF(TblSector,E{r},TblMV)')
        v.style = "td"; v.number_format = '"$"#,##0'
        a = ws.cell(row=r, column=7, value=f'=SUMIF(TblSector,E{r},TblIncome)')
        a.style = "td"; a.number_format = '"$"#,##0'
        w = ws.cell(row=r, column=8, value=f'=IFERROR(F{r}/TotalMarketValue,0)')
        w.style = "td"; w.number_format = "0.0%"
    sec_end = 6 + len(SECTORS)
    wb.defined_names["SectorLabels"] = DefinedName("SectorLabels", attr_text=f"Calculations!$E$7:$E${sec_end}")
    wb.defined_names["SectorValues"] = DefinedName("SectorValues", attr_text=f"Calculations!$F$7:$F${sec_end}")

    # Top dividend payers (ranked)
    top_row = sec_end + 2
    merge_set(ws, f"E{top_row}:H{top_row}", "TOP DIVIDEND PAYERS", "section_gold")
    hdr = top_row + 1
    for j, h in enumerate(["Rank", "Ticker", "Company", "Annual Inc"]):
        ws.cell(row=hdr, column=5 + j, value=h).style = "th"
    for k in range(1, 11):
        r = hdr + k
        ws.cell(row=r, column=5, value=k).style = "td"
        ws.cell(row=r, column=6,
                value=f'=IFERROR(INDEX(TblTicker,MATCH(LARGE(TblIncome,{k}),TblIncome,0)),"")').style = "td"
        ws.cell(row=r, column=7,
                value=f'=IFERROR(INDEX(TblCompany,MATCH(LARGE(TblIncome,{k}),TblIncome,0)),"")').style = "td_left"
        c = ws.cell(row=r, column=8, value=f'=IFERROR(LARGE(TblIncome,{k}),"")')
        c.style = "td"; c.number_format = '"$"#,##0'
    wb.defined_names["TopRank"] = DefinedName("TopRank", attr_text=f"Calculations!$E${hdr+1}:$E${hdr+10}")
    wb.defined_names["TopTicker"] = DefinedName("TopTicker", attr_text=f"Calculations!$F${hdr+1}:$F${hdr+10}")
    wb.defined_names["TopCompany"] = DefinedName("TopCompany", attr_text=f"Calculations!$G${hdr+1}:$G${hdr+10}")
    wb.defined_names["TopIncome"] = DefinedName("TopIncome", attr_text=f"Calculations!$H${hdr+1}:$H${hdr+10}")


# ===========================================================================
# Income Calendar
# ===========================================================================
def build_income_calendar(wb):
    ws = wb.create_sheet("IncomeCalendar")
    ws.sheet_view.showGridLines = False
    set_widths(ws, [10, 14, 14] + [9] * 12)
    luxe_header(ws, "O", "📅  DIVIDEND INCOME CALENDAR",
                "See exactly which months you get paid — built from each holding's frequency.")
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    ws.cell(row=4, column=1, value="Ticker").style = "th"
    ws.cell(row=4, column=2, value="Annual Inc").style = "th"
    ws.cell(row=4, column=3, value="Frequency").style = "th"
    for i, m in enumerate(months):
        ws.cell(row=4, column=4 + i, value=m).style = "th"
    ws.row_dimensions[4].height = 28

    for r in range(H_START, H_END + 1):
        ws.cell(row=r, column=1, value=f'=IF(Holdings!A{r}="","",Holdings!A{r})')
        ws.cell(row=r, column=2, value=f'=IF(Holdings!A{r}="","",Holdings!M{r})')
        ws.cell(row=r, column=3, value=f'=IF(Holdings!A{r}="","",Holdings!L{r})')
        for m in range(1, 13):
            col = 3 + m
            ws.cell(row=r, column=col, value=(
                f'=IF($A{r}="","",'
                f'IF($C{r}="Monthly",$B{r}/12,'
                f'IF(AND($C{r}="Quarterly",MOD({m},3)=0),$B{r}/4,'
                f'IF(AND($C{r}="Semi-Annual",OR({m}=6,{m}=12)),$B{r}/2,'
                f'IF(AND($C{r}="Annual",{m}=12),$B{r},0)))))'))

    money = {2} | set(range(4, 16))
    for r in range(H_START, H_END + 1):
        for c in range(1, 16):
            cell = ws.cell(row=r, column=c)
            cell.style = "td_left" if c == 1 else "td"
            cell.fill = fill(MUTED_ROW if (r - H_START) % 2 else BG)
            if c in money:
                cell.number_format = '"$"#,##0'
        ws.cell(row=r, column=1).font = Font(bold=True, color=PRIMARY)

    # Monthly totals
    tot = H_END + 1
    ws.cell(row=tot, column=1, value="TOTAL").style = "th"
    ws.merge_cells(start_row=tot, start_column=1, end_row=tot, end_column=3)
    ws.cell(row=tot, column=1).style = "th"
    for m in range(1, 13):
        col = 3 + m
        L = get_column_letter(col)
        c = ws.cell(row=tot, column=col, value=f"=SUM({L}{H_START}:{L}{H_END})")
        c.style = "td"; c.font = Font(bold=True, color=PRIMARY)
        c.fill = fill(SURFACE); c.number_format = '"$"#,##0'

    # Color-scale the month cells to make the calendar pop
    ws.conditional_formatting.add(
        f"D{H_START}:O{H_END}",
        ColorScaleRule(start_type="num", start_value=0, start_color="FFFFFFFF",
                       end_type="max", end_color="FF" + HIGHLIGHT))

    wb.defined_names["CalMonthLabels"] = DefinedName("CalMonthLabels", attr_text="IncomeCalendar!$D$4:$O$4")
    wb.defined_names["CalMonthTotals"] = DefinedName("CalMonthTotals", attr_text=f"IncomeCalendar!$D${tot}:$O${tot}")
    ws.freeze_panes = "D5"


# ===========================================================================
# Projections & FIRE
# ===========================================================================
def build_projections(wb):
    ws = wb.create_sheet("Projections")
    ws.sheet_view.showGridLines = False
    set_widths(ws, [2, 10, 20, 20, 18, 22, 14])
    luxe_header(ws, "G", "🚀  PROJECTIONS & FIRE MODEL",
                "Compounding forecast driven by your Settings inputs.")

    merge_set(ws, "B5:G5", "FINANCIAL INDEPENDENCE METRICS", "section_gold")
    fire = [
        ("FIRE Number (25× annual goal)", "=TargetMonthlyIncome*12*25", '"$"#,##0', "FireNumber"),
        ("Years to $1,000 / month",
         '=IFERROR(IF(TotalMonthlyIncome>=1000,0,LOG(1000/(TotalMonthlyIncome+0.0001))/LOG(1+DividendGrowthRate)),"—")',
         "0.0", "YearsTo1k"),
        ("Years to $5,000 / month",
         '=IFERROR(IF(TotalMonthlyIncome>=5000,0,LOG(5000/(TotalMonthlyIncome+0.0001))/LOG(1+DividendGrowthRate)),"—")',
         "0.0", "YearsTo5k"),
        ("Years to Target Income",
         '=IFERROR(IF(TotalMonthlyIncome>=TargetMonthlyIncome,0,LOG(TargetMonthlyIncome/(TotalMonthlyIncome+0.0001))/LOG(1+DividendGrowthRate)),"—")',
         "0.0", "YearsToTarget"),
    ]
    for i, (lab, fml, fmt, name) in enumerate(fire):
        r = 6 + i
        ws.cell(row=r, column=2, value=lab).style = "field_label"
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        ws.cell(row=r, column=2).style = "field_label"
        c = ws.cell(row=r, column=5, value=fml)
        c.style = "field_value"; c.number_format = fmt
        wb.defined_names[name] = DefinedName(name, attr_text=f"Projections!$E${r}")

    merge_set(ws, "B11:G11", "30-YEAR PROJECTION", "section_gold")
    hdrs = ["Year", "Portfolio Value", "Annual Income", "Monthly Income",
            "Cumulative Contributions", "% to FIRE"]
    for i, h in enumerate(hdrs, 2):
        ws.cell(row=12, column=i, value=h).style = "th"
    ws.row_dimensions[12].height = 30

    ps, pe = 13, 13 + 30 - 1
    for i in range(30):
        r = ps + i
        ws.cell(row=r, column=2, value=i + 1).style = "td"
        if i == 0:
            ws.cell(row=r, column=3, value="=TotalMarketValue*(1+AnnualGrowthRate)+MonthlyContribution*12")
            ws.cell(row=r, column=4, value="=TotalAnnualIncome*(1+DividendGrowthRate)")
            ws.cell(row=r, column=6, value="=MonthlyContribution*12")
        else:
            ws.cell(row=r, column=3,
                    value=f"=C{r-1}*(1+AnnualGrowthRate)+MonthlyContribution*12+IF(ReinvestToggle=1,D{r-1},0)")
            ws.cell(row=r, column=4,
                    value=f"=D{r-1}*(1+DividendGrowthRate)+IF(ReinvestToggle=1,C{r-1}*PortfolioYield*0.05,0)")
            ws.cell(row=r, column=6, value=f"=F{r-1}+MonthlyContribution*12")
        ws.cell(row=r, column=5, value=f"=D{r}/12")
        ws.cell(row=r, column=7, value=f"=IFERROR(C{r}/FireNumber,0)")
        for c in range(2, 8):
            cell = ws.cell(row=r, column=c)
            cell.style = "td"
            cell.fill = fill(MUTED_ROW if i % 2 else BG)
        for c in (3, 4, 5, 6):
            ws.cell(row=r, column=c).number_format = '"$"#,##0'
        ws.cell(row=r, column=7).number_format = "0.0%"
        ws.cell(row=r, column=2).number_format = "0"

    wb.defined_names["ProjYears"] = DefinedName("ProjYears", attr_text=f"Projections!$B${ps}:$B${pe}")
    wb.defined_names["ProjPortfolio"] = DefinedName("ProjPortfolio", attr_text=f"Projections!$C${ps}:$C${pe}")
    wb.defined_names["ProjAnnualIncome"] = DefinedName("ProjAnnualIncome", attr_text=f"Projections!$D${ps}:$D${pe}")

    # Charts (stacked on the right, clear of the table)
    wealth = LineChart(); wealth.title = "Projected Wealth Growth"; wealth.height = 8; wealth.width = 16
    wealth.add_data(Reference(ws, min_col=3, min_row=12, max_row=pe), titles_from_data=True)
    wealth.set_categories(Reference(ws, min_col=2, min_row=ps, max_row=pe))
    ws.add_chart(wealth, "I5")
    inc = LineChart(); inc.title = "Annual Dividend Income"; inc.height = 8; inc.width = 16
    inc.add_data(Reference(ws, min_col=4, min_row=12, max_row=pe), titles_from_data=True)
    inc.set_categories(Reference(ws, min_col=2, min_row=ps, max_row=pe))
    ws.add_chart(inc, "I22")


# ===========================================================================
# Dashboard
# ===========================================================================
def build_dashboard(wb):
    ws = wb.create_sheet("Dashboard", 0)
    ws.sheet_view.showGridLines = False
    set_widths(ws, [2] + [15] * 12 + [2])
    ws.row_dimensions[1].height = 56
    merge_set(ws, "A1:N1", "  💎  DIVIDEND WEALTH BUILDER", "title")
    ws.row_dimensions[2].height = 24
    merge_set(ws, "A2:N2",
              "  Track every dividend, optimize your yield, and forecast financial independence — in one elegant system.",
              "subtitle")
    ws.row_dimensions[3].height = 6
    for c in range(1, 15):
        ws.cell(row=3, column=c).fill = fill(GOLD_LT)

    # KPI cards — 4 per row, each 3 cols wide (B-D, E-G, H-J, K-M)
    row1 = [
        ("PORTFOLIO VALUE", "=TotalMarketValue", "money"),
        ("MONTHLY INCOME", "=TotalMonthlyIncome", "money2"),
        ("ANNUAL INCOME", "=TotalAnnualIncome", "money"),
        ("PORTFOLIO YIELD", "=PortfolioYield", "pct"),
    ]
    row2 = [
        ("YIELD ON COST", "=PortfolioYoC", "pct"),
        ("TOTAL GAIN / LOSS", "=TotalGain", "money"),
        ("DIVIDEND GROWTH", "=DividendGrowthRate", "pct"),
        ("PROGRESS TO GOAL", "=ProgressToGoal", "pct"),
    ]
    cols = [2, 5, 8, 11]
    for (lab, fml, kind), col in zip(row1, cols):
        kpi_card(ws, 5, col, 3, lab, fml, kind)
    for (lab, fml, kind), col in zip(row2, cols):
        kpi_card(ws, 8, col, 3, lab, fml, kind)

    # Analytics
    ws.row_dimensions[11].height = 26
    merge_set(ws, "B11:M11", "PORTFOLIO ANALYTICS", "section_gold")

    donut = DoughnutChart(); donut.title = "Allocation by Sector"; donut.height = 8.5; donut.width = 12
    donut.add_data(Reference(wb["Calculations"], min_col=6, min_row=6, max_row=6 + len(SECTORS)),
                   titles_from_data=True)
    donut.set_categories(Reference(wb["Calculations"], min_col=5, min_row=7, max_row=6 + len(SECTORS)))
    donut.dataLabels = clean_labels(pct=True)
    ws.add_chart(donut, "B12")

    bar = BarChart(); bar.type = "col"; bar.title = "Dividend Income by Month"; bar.height = 8.5; bar.width = 12
    bar.add_data(Reference(wb["IncomeCalendar"], min_col=4, min_row=H_END + 1,
                           max_col=15, max_row=H_END + 1), titles_from_data=False)
    bar.set_categories(Reference(wb["IncomeCalendar"], min_col=4, min_row=4, max_col=15, max_row=4))
    bar.legend = None
    ws.add_chart(bar, "H12")

    # Top payers table (left) + wealth line (right)
    ws.row_dimensions[30].height = 26
    merge_set(ws, "B30:G30", "TOP DIVIDEND PAYERS", "section")
    for j, h in enumerate(["#", "Ticker", "Company", "Annual Income"]):
        ws.cell(row=31, column=2 + j, value=h).style = "th"
    ws.merge_cells("D31:E31"); ws.cell(row=31, column=4).style = "th"; ws.cell(row=31, column=4).value = "Company"
    ws.cell(row=31, column=6).value = "Annual Income"; ws.cell(row=31, column=6).style = "th"
    for k in range(1, 9):
        r = 31 + k
        ws.cell(row=r, column=2, value=f"=IFERROR(INDEX(TopRank,{k}),\"\")").style = "td"
        ws.cell(row=r, column=3, value=f"=IFERROR(INDEX(TopTicker,{k}),\"\")").style = "td"
        ws.merge_cells(f"D{r}:E{r}")
        cc = ws.cell(row=r, column=4, value=f"=IFERROR(INDEX(TopCompany,{k}),\"\")")
        cc.style = "td_left"
        c = ws.cell(row=r, column=6, value=f"=IFERROR(INDEX(TopIncome,{k}),\"\")")
        c.style = "td"; c.number_format = '"$"#,##0'
        for cidx in (2, 3, 4, 5, 6):
            ws.cell(row=r, column=cidx).fill = fill(MUTED_ROW if k % 2 == 0 else BG)
            if ws.cell(row=r, column=cidx).border.left is None:
                ws.cell(row=r, column=cidx).border = BOX

    line = LineChart(); line.title = "Projected Wealth Growth"; line.height = 8.5; line.width = 12
    line.add_data(Reference(wb["Projections"], min_col=3, min_row=12, max_row=42), titles_from_data=True)
    line.set_categories(Reference(wb["Projections"], min_col=2, min_row=13, max_row=42))
    ws.add_chart(line, "H30")

    # Footer
    ws.row_dimensions[48].height = 26
    merge_set(ws, "B48:M48",
              "Dividend Wealth Builder v2.0  ·  Edit Settings → contribution · target · growth  ·  Add holdings on the Holdings tab",
              "subtitle")


# ===========================================================================
# Build
# ===========================================================================
def main():
    wb = Workbook()
    wb.remove(wb.active)
    register_styles(wb)

    build_settings(wb)
    build_holdings(wb)
    build_calculations(wb)
    build_income_calendar(wb)
    build_projections(wb)
    build_dashboard(wb)

    order = ["Dashboard", "Holdings", "IncomeCalendar", "Calculations", "Projections", "Settings"]
    wb._sheets = [wb[n] for n in order]
    colors = {"Dashboard": PRIMARY, "Holdings": ACCENT, "IncomeCalendar": HIGHLIGHT,
              "Calculations": SURFACE, "Projections": PRIMARY, "Settings": SURFACE}
    for n, col in colors.items():
        wb[n].sheet_properties.tabColor = col

    out = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                       "Dividend_Wealth_Builder.xlsx")
    wb.save(out)
    print(f"Wrote {out}  ({len(order)} sheets)")


# clean_labels helper (shared pattern from the other products)
def clean_labels(pct=False, val=False):
    dl = DataLabelList()
    dl.showSerName = False
    dl.showCatName = False
    dl.showLegendKey = False
    dl.showBubbleSize = False
    dl.showVal = val
    dl.showPercent = pct
    return dl


if __name__ == "__main__":
    main()

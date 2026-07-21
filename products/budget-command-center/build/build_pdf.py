"""Build the printable PDF pack for Budget Command Center™ (12 pages, US Letter).

  1  Monthly Budget (Zero-Based)  7  Net Worth Worksheet
  2  Bill Tracker                 8  Subscriptions Audit
  3  Expense Log                  9  Income Tracker
  4  Savings Goals               10  No-Spend Challenge
  5  Sinking Funds               11  Year-at-a-Glance
  6  Debt Snowball / Payoff      12  Money Goals

Outputs ../Budget_Printables.pdf and page PNGs in ../marketing/print/.
Run: python3 build_pdf.py
"""
from __future__ import annotations
import os
from PIL import Image, ImageDraw, ImageFont

PRIMARY = (27, 79, 72); PRIMARY_DK = (18, 56, 51); ACCENT = (147, 115, 86)
GOLD = (180, 145, 90); GOLD_LT = (201, 168, 106); GOLD_HI = (224, 196, 140)
SURFACE = (229, 211, 186); MINT = (227, 248, 239); WARN = (251, 240, 226)
WHITE = (255, 255, 255); TEXT = (51, 51, 51); MUTED = (120, 114, 104)
LINE = (200, 194, 182); ROW_ALT = (247, 243, 236)

W, H = 2550, 3300
M = 190

SANS_B = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
SANS_R = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
SERIF_B = "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf"


def fs(s, bold=True):
    return ImageFont.truetype(SANS_B if bold else SANS_R, s)


def fserif(s):
    return ImageFont.truetype(SERIF_B, s)


def page(title, subtitle):
    img = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(img)
    bh = 340
    d.rectangle((0, 0, W, bh), fill=PRIMARY)
    d.rectangle((0, bh, W, bh + 10), fill=GOLD_LT)
    d.rectangle((0, bh + 10, W, bh + 14), fill=GOLD_HI)
    d.text((M, 90), "BUDGET COMMAND CENTER™", font=fs(34), fill=GOLD_HI, anchor="lt")
    d.text((M, 150), title, font=fserif(74), fill=WHITE, anchor="lt")
    if subtitle:
        d.text((M, 258), subtitle, font=fs(30, bold=False), fill=(214, 226, 222), anchor="lt")
    d.line((M, H - 130, W - M, H - 130), fill=LINE, width=2)
    d.text((M, H - 108), "Month: ______________", font=fs(24, bold=False), fill=MUTED, anchor="lt")
    d.text((W - M, H - 108), "Budget Command Center™", font=fs(24, bold=False), fill=ACCENT, anchor="rt")
    return img, d


def checkbox(d, x, y, label, size=44, font=None, fill=TEXT):
    font = font or fs(30, bold=False)
    d.rounded_rectangle((x, y, x + size, y + size), radius=8, outline=PRIMARY, width=3)
    d.text((x + size + 24, y + size / 2), label, font=font, fill=fill, anchor="lm")


def field(d, x, y, w, label, lab_font=None, line=True):
    lab_font = lab_font or fs(24)
    d.text((x, y), label.upper(), font=lab_font, fill=ACCENT, anchor="lt")
    if line:
        d.line((x, y + 62, x + w, y + 62), fill=LINE, width=2)


def section(d, x, y, w, text):
    d.rounded_rectangle((x, y, x + w, y + 56), radius=8, fill=SURFACE)
    d.text((x + 20, y + 28), text.upper(), font=fs(26), fill=PRIMARY, anchor="lm")


def table(d, x, y, w, headers, colf, nrows, rowh=92, filled_rows=None):
    filled_rows = filled_rows or []
    colx = [x + w * f for f in colf]
    hh = 78
    d.rounded_rectangle((x, y, x + w, y + hh), radius=8, fill=PRIMARY)
    for i, h in enumerate(headers):
        d.text((colx[i] + (18 if i == 0 else 0), y + hh / 2), h, font=fs(26),
               fill=WHITE, anchor="lm" if i == 0 else "mm")
    for r in range(nrows):
        ry = y + hh + r * rowh
        if r % 2:
            d.rectangle((x, ry, x + w, ry + rowh), fill=ROW_ALT)
        d.line((x, ry + rowh, x + w, ry + rowh), fill=LINE, width=2)
        if r < len(filled_rows):
            for ci, val in enumerate(filled_rows[r]):
                d.text((colx[ci] + (18 if ci == 0 else 0), ry + rowh / 2), str(val),
                       font=fs(26, bold=(ci == 0)), fill=PRIMARY if ci == 0 else TEXT,
                       anchor="lm" if ci == 0 else "mm")
    d.rectangle((x, y, x + w, y + hh + nrows * rowh), outline=LINE, width=2)
    for cx in colx[1:]:
        d.line((cx, y, cx, y + hh + nrows * rowh), fill=LINE, width=1)


CW = W - 2 * M
TOP = 430


# ---------------------------------------------------------------- pages
def p01(imgs):
    img, d = page("Monthly Budget", "Give every dollar a job — planned vs actual")
    d.text((M, TOP), "INCOME THIS MONTH:", font=fs(26), fill=ACCENT, anchor="lt")
    d.line((M + 460, TOP + 34, M + 900, TOP + 34), fill=LINE, width=2)
    d.text((M + 1000, TOP), "LEFT TO BUDGET:", font=fs(26), fill=ACCENT, anchor="lt")
    d.line((M + 1380, TOP + 34, W - M, TOP + 34), fill=LINE, width=2)
    table(d, M, TOP + 90, CW, ["CATEGORY", "PLANNED", "ACTUAL", "REMAINING"],
          [0.0, 0.46, 0.64, 0.82], 18, rowh=130)
    imgs.append(img)


def p02(imgs):
    img, d = page("Bill Tracker", "Every recurring bill — never a late fee again")
    table(d, M, TOP, CW, ["BILL", "AMOUNT", "DUE DAY", "PAID?", "AUTOPAY?"],
          [0.0, 0.40, 0.58, 0.74, 0.88], 18, rowh=120)
    imgs.append(img)


def p03(imgs):
    img, d = page("Expense Log", "Every transaction — the truth about your spending")
    table(d, M, TOP, CW, ["DATE", "MERCHANT", "CATEGORY", "AMOUNT"],
          [0.0, 0.14, 0.50, 0.82], 20, rowh=118)
    imgs.append(img)


def p04(imgs):
    img, d = page("Savings Goals", "Color the bar as you save — a full bar is a great day")
    goals = ["Emergency Fund", "Vacation Fund", "New Car Fund", "Home Projects", "Holidays", "Big Purchase"]
    y = TOP
    for g in goals:
        section(d, M, y, CW, g); y += 80
        field(d, M, y, CW / 2 - 30, "Target"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Saved so far")
        y += 100
        d.rounded_rectangle((M, y, M + CW, y + 60), radius=14, outline=PRIMARY, width=3)
        for k in range(10):
            d.line((M + CW * (k + 1) / 10, y, M + CW * (k + 1) / 10, y + 60), fill=LINE, width=1)
        y += 130
    imgs.append(img)


def p05(imgs):
    img, d = page("Sinking Funds", "Save a little each month for the big irregular costs")
    table(d, M, TOP, CW, ["FUND", "MONTHLY", "GOAL", "BALANCE"],
          [0.0, 0.42, 0.60, 0.80], 16, rowh=130)
    imgs.append(img)


def p06(imgs):
    img, d = page("Debt Snowball / Payoff", "List smallest to largest — attack one at a time")
    table(d, M, TOP, CW, ["DEBT", "BALANCE", "APR", "MIN PAY", "PAID OFF?"],
          [0.0, 0.36, 0.54, 0.68, 0.86], 12, rowh=170)
    imgs.append(img)


def p07(imgs):
    img, d = page("Net Worth Worksheet", "Assets minus liabilities — the number that tells the truth")
    half = CW / 2 - 30
    section(d, M, TOP, half, "Assets (what you own)")
    section(d, M + CW / 2 + 30, TOP, half, "Liabilities (what you owe)")
    for i in range(9):
        y = TOP + 96 + i * 118
        d.text((M, y), "•", font=fs(30), fill=GOLD_LT, anchor="lt")
        d.line((M + 44, y + 44, M + half, y + 44), fill=LINE, width=2)
        d.text((M + CW / 2 + 30, y), "•", font=fs(30), fill=GOLD_LT, anchor="lt")
        d.line((M + CW / 2 + 74, y + 44, W - M, y + 44), fill=LINE, width=2)
    y = TOP + 96 + 9 * 118 + 20
    field(d, M, y, half, "Total assets"); field(d, M + CW / 2 + 30, y, half, "Total liabilities")
    y += 130
    section(d, M, y, CW, "Net worth  =  total assets  −  total liabilities  =  $ __________")
    imgs.append(img)


def p08(imgs):
    img, d = page("Subscriptions Audit", "The quiet budget-killer — cancel what you don't use")
    table(d, M, TOP, CW, ["SUBSCRIPTION", "MONTHLY", "ANNUAL", "KEEP?"],
          [0.0, 0.42, 0.60, 0.82], 18, rowh=120)
    imgs.append(img)


def p09(imgs):
    img, d = page("Income Tracker", "Every paycheck & source — the top of your budget")
    table(d, M, TOP, CW, ["DATE", "SOURCE", "GROSS", "NET"],
          [0.0, 0.16, 0.56, 0.78], 16, rowh=130)
    imgs.append(img)


def p10(imgs):
    img, d = page("No-Spend Challenge", "Color a day green for every no-spend day")
    cols, rows = 7, 5
    gx, gy = M + 200, TOP + 60
    bw = (CW - 200) / cols
    days = ["S", "M", "T", "W", "T", "F", "S"]
    for i, dn in enumerate(days):
        d.text((gx + bw * i + bw / 2, TOP + 20), dn, font=fs(28), fill=PRIMARY, anchor="mm")
    n = 1
    for r in range(rows):
        for c in range(cols):
            x = gx + bw * c; y = gy + r * bw
            d.rounded_rectangle((x + 6, y + 6, x + bw - 6, y + bw - 6), radius=12, outline=PRIMARY, width=3)
            if n <= 31:
                d.text((x + 24, y + 24), str(n), font=fs(24, bold=False), fill=MUTED, anchor="lt")
            n += 1
    y = gy + rows * bw + 40
    field(d, M, y, CW / 2 - 30, "No-spend days this month")
    field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Money saved")
    imgs.append(img)


def p11(imgs):
    img, d = page("Year-at-a-Glance", "The whole year, month by month")
    table(d, M, TOP, CW, ["MONTH", "INCOME", "SPENT", "SAVED", "NET WORTH"],
          [0.0, 0.28, 0.46, 0.64, 0.82], 12, rowh=170,
          filled_rows=[[m] for m in ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]])
    imgs.append(img)


def p12(imgs):
    img, d = page("Money Goals", "The big picture — where you're headed & why")
    y = TOP
    for who in ["This month", "This year", "3–5 years", "Why it matters to me"]:
        section(d, M, y, CW, who); y += 96
        for i in range(3):
            checkbox(d, M, y + i * 84, "", size=40)
            d.line((M + 66, y + i * 84 + 40, M + CW, y + i * 84 + 40), fill=LINE, width=2)
        y += 3 * 84 + 40
    imgs.append(img)


def main():
    out_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    print_dir = os.path.join(out_dir, "marketing", "print")
    os.makedirs(print_dir, exist_ok=True)
    imgs = []
    for fn in [p01, p02, p03, p04, p05, p06, p07, p08, p09, p10, p11, p12]:
        fn(imgs)
    for i, im in enumerate(imgs, 1):
        im.save(os.path.join(print_dir, f"page_{i:02d}.png"), "PNG", optimize=True)
    pdf_path = os.path.join(out_dir, "Budget_Printables.pdf")
    imgs[0].save(pdf_path, "PDF", resolution=300.0, save_all=True, append_images=imgs[1:])
    print(f"Wrote {pdf_path}  ({len(imgs)} pages) + page PNGs in {print_dir}")


if __name__ == "__main__":
    main()

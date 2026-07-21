"""Build the printable PDF pack for Debt Payoff Command Center™ (12 pages, US Letter).

  1  Debt List & Snapshot          7  Interest Saved Worksheet
  2  Payoff Plan (Attack Order)    8  Extra-Payment Finder
  3  Snowball vs Avalanche         9  Payment Log
  4  Debt Payoff Tracker (per debt)10 Balance History Chart
  5  Debt-Free Date Worksheet     11  Milestones & Wins
  6  Debt Thermometer             12  Debt-Freedom Goals

Outputs ../Debt_Printables.pdf and page PNGs in ../marketing/print/.
Run: python3 build_pdf.py
"""
from __future__ import annotations
import os
from PIL import Image, ImageDraw, ImageFont

PRIMARY = (27, 79, 72); PRIMARY_DK = (18, 56, 51); ACCENT = (147, 115, 86)
GOLD = (180, 145, 90); GOLD_LT = (201, 168, 106); GOLD_HI = (224, 196, 140)
SURFACE = (229, 211, 186); MINT = (227, 248, 239); WARN = (251, 240, 226)
WHITE = (255, 255, 255); TEXT = (51, 51, 51); MUTED = (120, 114, 104)
LINE = (200, 194, 182); ROW_ALT = (247, 243, 236); DANGER = (201, 76, 76)

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
    d.text((M, 90), "DEBT PAYOFF COMMAND CENTER™", font=fs(34), fill=GOLD_HI, anchor="lt")
    d.text((M, 150), title, font=fserif(74), fill=WHITE, anchor="lt")
    if subtitle:
        d.text((M, 258), subtitle, font=fs(30, bold=False), fill=(214, 226, 222), anchor="lt")
    d.line((M, H - 130, W - M, H - 130), fill=LINE, width=2)
    d.text((M, H - 108), "Date: ______________", font=fs(24, bold=False), fill=MUTED, anchor="lt")
    d.text((W - M, H - 108), "Debt Payoff Command Center™", font=fs(24, bold=False), fill=ACCENT, anchor="rt")
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
    img, d = page("Debt List & Snapshot", "List every debt once — this is the map")
    table(d, M, TOP, CW, ["DEBT", "BALANCE", "APR", "MIN PAY", "TYPE"],
          [0.0, 0.36, 0.54, 0.68, 0.85], 14, rowh=150)
    y = TOP + 78 + 14 * 150 + 30
    field(d, M, y, CW / 2 - 30, "Total balance")
    field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Total minimum payments")
    imgs.append(img)


def p02(imgs):
    img, d = page("Payoff Plan", "Number them in attack order — one focus debt at a time")
    table(d, M, TOP, CW, ["ORDER", "DEBT", "BALANCE", "EXTRA $", "TARGET DATE"],
          [0.0, 0.16, 0.46, 0.64, 0.82], 12, rowh=170)
    imgs.append(img)


def p03(imgs):
    img, d = page("Snowball vs Avalanche", "Two methods — pick the one you'll actually stick to")
    half = CW / 2 - 30
    section(d, M, TOP, half, "Snowball — smallest balance first")
    section(d, M + CW / 2 + 30, TOP, half, "Avalanche — highest rate first")
    for i in range(9):
        y = TOP + 96 + i * 130
        d.text((M, y), f"{i+1}.", font=fs(30), fill=GOLD_LT, anchor="lt")
        d.line((M + 80, y + 44, M + half, y + 44), fill=LINE, width=2)
        d.text((M + CW / 2 + 30, y), f"{i+1}.", font=fs(30), fill=GOLD_LT, anchor="lt")
        d.line((M + CW / 2 + 110, y + 44, W - M, y + 44), fill=LINE, width=2)
    y = TOP + 96 + 9 * 130 + 20
    field(d, M, y, half, "Fastest first win (snowball)")
    field(d, M + CW / 2 + 30, y, half, "Interest saved (avalanche)")
    imgs.append(img)


def p04(imgs):
    img, d = page("Debt Payoff Tracker", "Color a box for every $100 you knock out")
    debts = ["Store Card", "Credit Card A", "Personal Loan", "Car Loan", "Student Loan", "Other"]
    y = TOP
    for g in debts:
        section(d, M, y, CW, g); y += 76
        field(d, M, y, CW / 2 - 30, "Starting balance")
        field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Paid off so far")
        y += 96
        d.rounded_rectangle((M, y, M + CW, y + 56), radius=14, outline=PRIMARY, width=3)
        for k in range(20):
            d.line((M + CW * (k + 1) / 20, y, M + CW * (k + 1) / 20, y + 56), fill=LINE, width=1)
        y += 120
    imgs.append(img)


def p05(imgs):
    img, d = page("Debt-Free Date Worksheet", "Do the math — then circle the month you're free")
    rows = [
        "Total debt balance:  $ ______________",
        "Total minimum payments / month:  $ ____________",
        "Extra payment / month:  $ ____________",
        "Total monthly payment (min + extra):  $ ___________",
        "Average interest rate (APR):  __________ %",
        "Estimated months to debt-free:  ____________",
        "My debt-free date:  ______________________",
    ]
    y = TOP + 20
    for r in rows:
        d.text((M, y), r, font=fs(34, bold=False), fill=TEXT, anchor="lt")
        y += 150
    y += 20
    section(d, M, y, CW, "The one thing I'll do this month to get there faster")
    y += 96
    d.line((M, y + 40, W - M, y + 40), fill=LINE, width=2)
    d.line((M, y + 150, W - M, y + 150), fill=LINE, width=2)
    imgs.append(img)


def p06(imgs):
    img, d = page("Debt Thermometer", "Color it up as you pay down — from total to $0")
    cx = W // 2 - 60
    top = TOP + 80
    tube_w = 220
    R = 250                       # bulb radius
    bulb_cy = H - 520
    tube_bot = bulb_cy - R // 2
    # tube (round only matters at top; bottom is hidden by the bulb)
    d.rounded_rectangle((cx - tube_w // 2, top, cx + tube_w // 2, tube_bot + 80),
                        radius=tube_w // 2, outline=PRIMARY, width=6)
    # bulb — white fill hides the tube's lower arc, then its own outline
    d.ellipse((cx - R, bulb_cy - R, cx + R, bulb_cy + R), fill=WHITE, outline=PRIMARY, width=6)
    d.text((cx, bulb_cy), "$0", font=fserif(72), fill=PRIMARY, anchor="mm")
    # scale ticks up the right side of the tube
    n = 10
    seg = (tube_bot - top) / n
    for i in range(n + 1):
        yy = top + i * seg
        d.line((cx + tube_w // 2 + 20, yy, cx + tube_w // 2 + 70, yy), fill=LINE, width=3)
        pct = 100 - i * 10
        d.text((cx + tube_w // 2 + 90, yy), f"{pct}% left", font=fs(30, bold=False), fill=MUTED, anchor="lm")
    d.text((cx - tube_w // 2 - 40, top), "START", font=fs(28), fill=DANGER, anchor="rm")
    d.text((cx - tube_w // 2 - 40, tube_bot), "PAID!", font=fs(28), fill=PRIMARY, anchor="rm")
    imgs.append(img)


def p07(imgs):
    img, d = page("Interest Saved Worksheet", "See what interest costs — and what you'll save")
    rows = [
        "If I pay only minimums, interest costs:  $ ___________",
        "With my extra payment, interest costs:  $ ___________",
        "Extra payment SAVES me:  $ ______________",
        "",
        "Snowball total interest:  $ ______________",
        "Avalanche total interest:  $ ______________",
        "Avalanche saves vs snowball:  $ ___________",
    ]
    y = TOP + 20
    for r in rows:
        if r:
            d.text((M, y), r, font=fs(34, bold=False), fill=TEXT, anchor="lt")
        y += 150
    y += 10
    section(d, M, y, CW, "Every extra dollar is interest I never pay")
    imgs.append(img)


def p08(imgs):
    img, d = page("Extra-Payment Finder", "Hunt down money to throw at your focus debt")
    table(d, M, TOP, CW, ["SOURCE OF EXTRA MONEY", "MONTHLY $", "DONE?"],
          [0.0, 0.60, 0.84], 16, rowh=130)
    y = TOP + 78 + 16 * 130 + 30
    field(d, M, y, CW, "Total extra I can find each month")
    imgs.append(img)


def p09(imgs):
    img, d = page("Payment Log", "Every payment, logged — proof you're winning")
    table(d, M, TOP, CW, ["DATE", "DEBT", "AMOUNT", "NEW BALANCE"],
          [0.0, 0.16, 0.52, 0.76], 20, rowh=118)
    imgs.append(img)


def p10(imgs):
    img, d = page("Balance History", "Plot your total balance — watch the line fall")
    gx0, gy0 = M + 120, TOP + 40
    gx1, gy1 = W - M, H - 520
    d.line((gx0, gy0, gx0, gy1), fill=PRIMARY, width=4)
    d.line((gx0, gy1, gx1, gy1), fill=PRIMARY, width=4)
    for i in range(6):
        yy = gy0 + (gy1 - gy0) * i / 5
        d.line((gx0 - 20, yy, gx1, yy), fill=LINE, width=1)
        d.text((gx0 - 40, yy), f"{100 - i*20}%", font=fs(24, bold=False), fill=MUTED, anchor="rm")
    months = ["M1", "M2", "M3", "M4", "M5", "M6", "M7", "M8", "M9", "M10", "M11", "M12"]
    for i, m in enumerate(months):
        xx = gx0 + (gx1 - gx0) * (i + 0.5) / len(months)
        d.text((xx, gy1 + 30), m, font=fs(24, bold=False), fill=MUTED, anchor="mm")
    y = gy1 + 130
    field(d, M, y, CW / 2 - 30, "Starting balance")
    field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Balance now")
    imgs.append(img)


def p11(imgs):
    img, d = page("Milestones & Wins", "Debt payoff is a long road — celebrate every step")
    milestones = ["$1,000 starter emergency fund", "First debt paid off", "25% of total debt gone",
                  "Highest-rate debt cleared", "Half of all debt paid", "Down to one debt",
                  "Final payment made", "DEBT FREE 🎉"]
    y = TOP + 20
    for mstone in milestones:
        checkbox(d, M, y, mstone, size=54, font=fs(36, bold=False), fill=TEXT)
        d.line((M, y + 110, W - M, y + 110), fill=LINE, width=1)
        y += 168
    imgs.append(img)


def p12(imgs):
    img, d = page("Debt-Freedom Goals", "Why you're doing this — the reason on the hard days")
    y = TOP
    for who in ["When I'm debt-free I will…", "This debt is costing me…", "My reward at $0", "Why it matters to me"]:
        section(d, M, y, CW, who); y += 96
        for i in range(3):
            d.line((M, y + i * 84 + 40, M + CW, y + i * 84 + 40), fill=LINE, width=2)
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
    pdf_path = os.path.join(out_dir, "Debt_Printables.pdf")
    imgs[0].save(pdf_path, "PDF", resolution=300.0, save_all=True, append_images=imgs[1:])
    print(f"Wrote {pdf_path}  ({len(imgs)} pages) + page PNGs in {print_dir}")


if __name__ == "__main__":
    main()

"""Printable PDF pack for Savings Challenge & Sinking Funds Command Center™ (12 pages, US Letter)."""
from __future__ import annotations
import os
from PIL import Image, ImageDraw, ImageFont

PRIMARY = (27, 79, 72); ACCENT = (147, 115, 86); GOLD_LT = (201, 168, 106); GOLD_HI = (224, 196, 140)
SURFACE = (229, 211, 186); WHITE = (255, 255, 255); TEXT = (51, 51, 51); MUTED = (120, 114, 104)
LINE = (200, 194, 182); ROW_ALT = (247, 243, 236); MINT = (227, 248, 239)
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
    img = Image.new("RGB", (W, H), WHITE); d = ImageDraw.Draw(img)
    bh = 340
    d.rectangle((0, 0, W, bh), fill=PRIMARY)
    d.rectangle((0, bh, W, bh + 10), fill=GOLD_LT); d.rectangle((0, bh + 10, W, bh + 14), fill=GOLD_HI)
    d.text((M, 90), "SAVINGS CHALLENGE & SINKING FUNDS COMMAND CENTER™", font=fs(26), fill=GOLD_HI, anchor="lt")
    d.text((M, 150), title, font=fserif(74), fill=WHITE, anchor="lt")
    if subtitle:
        d.text((M, 258), subtitle, font=fs(30, bold=False), fill=(214, 226, 222), anchor="lt")
    d.line((M, H - 130, W - M, H - 130), fill=LINE, width=2)
    d.text((M, H - 108), "Year / Month: ______________", font=fs(24, bold=False), fill=MUTED, anchor="lt")
    d.text((W - M, H - 108), "Savings Challenge Command Center™", font=fs(22, bold=False), fill=ACCENT, anchor="rt")
    return img, d


def field(d, x, y, w, label):
    d.text((x, y), label.upper(), font=fs(24), fill=ACCENT, anchor="lt")
    d.line((x, y + 62, x + w, y + 62), fill=LINE, width=2)


def section(d, x, y, w, text):
    d.rounded_rectangle((x, y, x + w, y + 56), radius=8, fill=SURFACE)
    d.text((x + 20, y + 28), text.upper(), font=fs(26), fill=PRIMARY, anchor="lm")


def checkbox(d, x, y, size=48):
    d.rounded_rectangle((x, y, x + size, y + size), radius=8, outline=PRIMARY, width=3)


def table(d, x, y, w, headers, colf, nrows, rowh=92):
    colx = [x + w * f for f in colf]; hh = 78
    d.rounded_rectangle((x, y, x + w, y + hh), radius=8, fill=PRIMARY)
    for i, h in enumerate(headers):
        d.text((colx[i] + (18 if i == 0 else 0), y + hh / 2), h, font=fs(26), fill=WHITE, anchor="lm" if i == 0 else "mm")
    for r in range(nrows):
        ry = y + hh + r * rowh
        if r % 2:
            d.rectangle((x, ry, x + w, ry + rowh), fill=ROW_ALT)
        d.line((x, ry + rowh, x + w, ry + rowh), fill=LINE, width=2)
    d.rectangle((x, y, x + w, y + hh + nrows * rowh), outline=LINE, width=2)
    for cx in colx[1:]:
        d.line((cx, y, cx, y + hh + nrows * rowh), fill=LINE, width=1)


def grid_numbers(d, x, y, w, cols, rows, start_n=1, cell_h=None):
    """Numbered grid — the 100-envelope / 52-week colouring sheet."""
    cw = w / cols
    ch = cell_h or cw
    n = start_n
    for r in range(rows):
        for c in range(cols):
            cx0 = x + c * cw; cy0 = y + r * ch
            d.rectangle((cx0, cy0, cx0 + cw, cy0 + ch), outline=LINE, width=2)
            d.text((cx0 + cw / 2, cy0 + ch * 0.32), str(n), font=fs(int(ch * 0.26)), fill=PRIMARY, anchor="mm")
            d.text((cx0 + cw / 2, cy0 + ch * 0.72), f"${n}", font=fs(int(ch * 0.20), bold=False), fill=MUTED, anchor="mm")
            n += 1
    return n


CW = W - 2 * M; TOP = 430


def p01(i):
    img, d = page("Sinking Funds Worksheet", "Every bill that isn't monthly")
    table(d, M, TOP, CW, ["FUND", "TARGET", "SAVED", "MONTHLY SET-ASIDE"], [0.0, 0.42, 0.60, 0.76], 14, 165)
    y = TOP + 78 + 14 * 165 + 24
    field(d, M, y, CW / 2 - 30, "Total targets"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Total monthly set-aside")
    i.append(img)


def p02(i):
    img, d = page("100 Envelope Challenge", "Fill all 100 and you've saved $5,050")
    grid_numbers(d, M, TOP, CW, 10, 10, 1, cell_h=CW / 10 * 0.92)
    y = TOP + (CW / 10 * 0.92) * 10 + 40
    field(d, M, y, CW / 2 - 30, "Envelopes filled"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Total saved")
    i.append(img)


def p03(i):
    img, d = page("52-Week Challenge", "$1 in week 1 … $52 in week 52 = $1,378")
    grid_numbers(d, M, TOP, CW, 8, 7, 1, cell_h=CW / 8 * 0.95)
    y = TOP + (CW / 8 * 0.95) * 7 + 40
    field(d, M, y, CW / 2 - 30, "Weeks completed"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Total saved")
    i.append(img)


def p04(i):
    img, d = page("Savings Thermometer", "Colour it in as you climb")
    bx0 = M + CW * 0.30; bx1 = M + CW * 0.70
    top = TOP + 40; bot = H - 420
    d.rounded_rectangle((bx0, top, bx1, bot), radius=60, outline=PRIMARY, width=6)
    steps = 20
    seg = (bot - top) / steps
    for k in range(steps):
        yy = top + k * seg
        d.line((bx0, yy, bx1, yy), fill=LINE, width=2)
        d.text((bx0 - 24, yy + seg / 2), f"{100 - k * 5}%", font=fs(24, bold=False), fill=MUTED, anchor="rm")
    d.ellipse((bx0 - 40, bot - 40, bx1 + 40, bot + 120), fill=SURFACE, outline=PRIMARY, width=6)
    field(d, M, TOP - 90, CW / 2 - 30, "Goal")
    field(d, M + CW / 2 + 30, TOP - 90, CW / 2 - 30, "Target amount")
    i.append(img)


def p05(i):
    img, d = page("Cash Envelope Tracker", "Loaded, spent, left")
    table(d, M, TOP, CW, ["ENVELOPE", "LOADED", "SPENT", "LEFT"], [0.0, 0.40, 0.58, 0.78], 16, 148)
    i.append(img)


def p06(i):
    img, d = page("Deposit Log", "Proof the habit is real")
    table(d, M, TOP, CW, ["DATE", "INTO", "AMOUNT", "RUNNING TOTAL"], [0.0, 0.18, 0.48, 0.74], 20, 116)
    i.append(img)


def p07(i):
    img, d = page("Goal Countdown", "What you're saving toward")
    table(d, M, TOP, CW, ["GOAL", "TARGET", "SAVED", "TO GO", "BY WHEN"], [0.0, 0.34, 0.50, 0.66, 0.84], 14, 165)
    i.append(img)


def p08(i):
    img, d = page("No-Spend Tracker", "Colour every no-spend day")
    grid_numbers(d, M, TOP, CW, 7, 5, 1, cell_h=CW / 7 * 0.86)
    y = TOP + (CW / 7 * 0.86) * 5 + 50
    field(d, M, y, CW / 2 - 30, "No-spend days"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Saved this month")
    section(d, M, y + 160, CW, "What counts as a no-spend day")
    for j in range(4):
        d.line((M + 20, y + 300 + j * 90, W - M - 20, y + 300 + j * 90), fill=LINE, width=2)
    i.append(img)


def p09(i):
    img, d = page("Savings Streak", "31 days, one square each")
    grid_numbers(d, M, TOP, CW, 7, 5, 1, cell_h=CW / 7 * 0.86)
    y = TOP + (CW / 7 * 0.86) * 5 + 50
    field(d, M, y, CW / 2 - 30, "Current streak"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Longest streak")
    i.append(img)


def p10(i):
    img, d = page("Emergency Fund", "The one that catches you")
    field(d, M, TOP, CW / 2 - 30, "Goal amount"); field(d, M + CW / 2 + 30, TOP, CW / 2 - 30, "Saved so far")
    y = TOP + 160
    section(d, M, y, CW, "Monthly contributions")
    table(d, M, y + 90, CW, ["MONTH", "ADDED", "BALANCE"], [0.0, 0.42, 0.70], 12, 165)
    i.append(img)


def p11(i):
    img, d = page("Monthly Summary", "The line that keeps climbing")
    table(d, M, TOP, CW, ["MONTH", "SAVED", "CHALLENGES", "RUNNING TOTAL"], [0.0, 0.28, 0.50, 0.74], 14, 165)
    i.append(img)


def p12(i):
    img, d = page("Savings Checklist", "Set it up once")
    half = CW / 2 - 30
    for k, st in enumerate(["Set up your funds", "Every single month"]):
        x = M + k * (half + 60)
        section(d, x, TOP, half, st)
        for j in range(14):
            checkbox(d, x, TOP + 96 + j * 96)
            d.line((x + 72, TOP + 96 + j * 96 + 48, x + half, TOP + 96 + j * 96 + 48), fill=LINE, width=2)
    i.append(img)


def main():
    out_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    print_dir = os.path.join(out_dir, "marketing", "print"); os.makedirs(print_dir, exist_ok=True)
    imgs = []
    for fn in [p01, p02, p03, p04, p05, p06, p07, p08, p09, p10, p11, p12]:
        fn(imgs)
    for i, im in enumerate(imgs, 1):
        im.save(os.path.join(print_dir, f"page_{i:02d}.png"), "PNG", optimize=True)
    pdf_path = os.path.join(out_dir, "Savings_Challenge_Printables.pdf")
    imgs[0].save(pdf_path, "PDF", resolution=300.0, save_all=True, append_images=imgs[1:])
    print(f"Wrote {pdf_path}  ({len(imgs)} pages) + page PNGs in {print_dir}")


if __name__ == "__main__":
    main()

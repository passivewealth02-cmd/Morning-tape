"""Printable PDF pack for Contractor Job Costing & Bidding Command Center™ (12 pages, US Letter)."""
from __future__ import annotations
import os
from PIL import Image, ImageDraw, ImageFont

PRIMARY = (27, 79, 72); ACCENT = (147, 115, 86); GOLD_LT = (201, 168, 106); GOLD_HI = (224, 196, 140)
SURFACE = (229, 211, 186); WHITE = (255, 255, 255); TEXT = (51, 51, 51); MUTED = (120, 114, 104)
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
    img = Image.new("RGB", (W, H), WHITE); d = ImageDraw.Draw(img)
    bh = 340
    d.rectangle((0, 0, W, bh), fill=PRIMARY)
    d.rectangle((0, bh, W, bh + 10), fill=GOLD_LT); d.rectangle((0, bh + 10, W, bh + 14), fill=GOLD_HI)
    d.text((M, 90), "CONTRACTOR JOB COSTING & BIDDING COMMAND CENTER™", font=fs(26), fill=GOLD_HI, anchor="lt")
    d.text((M, 150), title, font=fserif(74), fill=WHITE, anchor="lt")
    if subtitle:
        d.text((M, 258), subtitle, font=fs(30, bold=False), fill=(214, 226, 222), anchor="lt")
    d.line((M, H - 130, W - M, H - 130), fill=LINE, width=2)
    d.text((M, H - 108), "Job / Date: ______________", font=fs(24, bold=False), fill=MUTED, anchor="lt")
    d.text((W - M, H - 108), "Contractor Command Center™", font=fs(22, bold=False), fill=ACCENT, anchor="rt")
    return img, d


def field(d, x, y, w, label):
    d.text((x, y), label.upper(), font=fs(24), fill=ACCENT, anchor="lt")
    d.line((x, y + 62, x + w, y + 62), fill=LINE, width=2)


def section(d, x, y, w, text):
    d.rounded_rectangle((x, y, x + w, y + 56), radius=8, fill=SURFACE)
    d.text((x + 20, y + 28), text.upper(), font=fs(26), fill=PRIMARY, anchor="lm")


def checkbox(d, x, y, size=48):
    d.rounded_rectangle((x, y, x + size, y + size), radius=8, outline=PRIMARY, width=3)


def table(d, x, y, w, headers, colf, nrows, rowh=92, filled_rows=None):
    filled_rows = filled_rows or []
    colx = [x + w * f for f in colf]; hh = 78
    d.rounded_rectangle((x, y, x + w, y + hh), radius=8, fill=PRIMARY)
    for i, h in enumerate(headers):
        d.text((colx[i] + (18 if i == 0 else 0), y + hh / 2), h, font=fs(26), fill=WHITE, anchor="lm" if i == 0 else "mm")
    for r in range(nrows):
        ry = y + hh + r * rowh
        if r % 2:
            d.rectangle((x, ry, x + w, ry + rowh), fill=ROW_ALT)
        d.line((x, ry + rowh, x + w, ry + rowh), fill=LINE, width=2)
        if r < len(filled_rows):
            for ci, val in enumerate(filled_rows[r]):
                d.text((colx[ci] + (18 if ci == 0 else 0), ry + rowh / 2), str(val), font=fs(26, bold=(ci == 0)),
                       fill=PRIMARY if ci == 0 else TEXT, anchor="lm" if ci == 0 else "mm")
    d.rectangle((x, y, x + w, y + hh + nrows * rowh), outline=LINE, width=2)
    for cx in colx[1:]:
        d.line((cx, y, cx, y + hh + nrows * rowh), fill=LINE, width=1)


CW = W - 2 * M; TOP = 430


def p01(i):
    img, d = page("Bid Worksheet", "The price you must quote")
    field(d, M, TOP, CW, "Job / Client")
    table(d, M, TOP + 130, CW, ["COST LINE", "AMOUNT"], [0.0, 0.78], 8, 150)
    y = TOP + 130 + 78 + 8 * 150 + 24
    field(d, M, y, CW / 3 - 20, "Direct cost"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "+ Overhead %")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "= Total cost")
    y2 = y + 150
    field(d, M, y2, CW / 2 - 30, "\u00f7 (1 \u2212 margin) = BID"); field(d, M + CW / 2 + 30, y2, CW / 2 - 30, "Planned profit")
    i.append(img)


def p02(i):
    img, d = page("Job Cost Sheet", "Estimate vs actual")
    table(d, M, TOP, CW, ["LINE", "ESTIMATED", "ACTUAL", "VARIANCE"], [0.0, 0.40, 0.60, 0.80], 14, 165)
    y = TOP + 78 + 14 * 165 + 24
    field(d, M, y, CW / 2 - 30, "Actual profit"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Actual margin %")
    i.append(img)


def p03(i):
    img, d = page("Materials Takeoff", "Every line, counted")
    table(d, M, TOP, CW, ["ITEM", "QTY", "UNIT COST", "TOTAL"], [0.0, 0.44, 0.60, 0.80], 18, 128)
    i.append(img)


def p04(i):
    img, d = page("Labor & Crew Log", "The biggest untracked cost")
    table(d, M, TOP, CW, ["NAME", "TRADE", "$ / HR", "HOURS", "COST"], [0.0, 0.30, 0.50, 0.66, 0.84], 16, 148)
    i.append(img)


def p05(i):
    img, d = page("Subcontractor Log", "Quoted vs invoiced")
    table(d, M, TOP, CW, ["SUB", "TRADE", "QUOTED", "INVOICED", "VARIANCE"], [0.0, 0.30, 0.48, 0.66, 0.84], 16, 148)
    i.append(img)


def p06(i):
    img, d = page("Change Order Form", "No signature, no work")
    field(d, M, TOP, CW / 2 - 30, "Job"); field(d, M + CW / 2 + 30, TOP, CW / 2 - 30, "Date")
    section(d, M, TOP + 150, CW, "Description of change")
    for j in range(6):
        d.line((M + 20, TOP + 290 + j * 90, W - M - 20, TOP + 290 + j * 90), fill=LINE, width=2)
    y = TOP + 290 + 6 * 90 + 40
    field(d, M, y, CW / 3 - 20, "Added cost"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "Added days")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "New contract total")
    y2 = y + 200
    field(d, M, y2, CW / 2 - 30, "Client signature"); field(d, M + CW / 2 + 30, y2, CW / 2 - 30, "Contractor signature")
    i.append(img)


def p07(i):
    img, d = page("Daily Job Log", "What happened on site")
    field(d, M, TOP, CW / 2 - 30, "Job"); field(d, M + CW / 2 + 30, TOP, CW / 2 - 30, "Date / Weather")
    table(d, M, TOP + 150, CW, ["CREW ON SITE", "HOURS", "WORK COMPLETED"], [0.0, 0.30, 0.46], 16, 148)
    i.append(img)


def p08(i):
    img, d = page("Bid Log", "Your win rate, tracked")
    table(d, M, TOP, CW, ["JOB", "BID AMOUNT", "RESULT", "WHY"], [0.0, 0.36, 0.54, 0.72], 18, 128)
    i.append(img)


def p09(i):
    img, d = page("Invoice & Draw Schedule", "Cash flow kills contractors")
    table(d, M, TOP, CW, ["JOB", "MILESTONE", "AMOUNT", "STATUS"], [0.0, 0.32, 0.54, 0.78], 18, 128)
    i.append(img)


def p10(i):
    img, d = page("Punch List", "Finish it properly")
    half = CW / 2 - 30
    for k, st in enumerate(["Interior", "Exterior & final"]):
        x = M + k * (half + 60)
        section(d, x, TOP, half, st)
        for j in range(16):
            checkbox(d, x, TOP + 96 + j * 92)
            d.line((x + 72, TOP + 96 + j * 92 + 48, x + half, TOP + 96 + j * 92 + 48), fill=LINE, width=2)
    i.append(img)


def p11(i):
    img, d = page("Monthly Summary", "Is the year working?")
    table(d, M, TOP, CW, ["MONTH", "REVENUE", "COSTS", "NET PROFIT"], [0.0, 0.28, 0.50, 0.76], 14, 165)
    i.append(img)


def p12(i):
    img, d = page("Job Closeout", "Before you invoice the last draw")
    half = CW / 2 - 30
    for k, st in enumerate(["Before closeout", "Handover to client"]):
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
    pdf_path = os.path.join(out_dir, "Contractor_Printables.pdf")
    imgs[0].save(pdf_path, "PDF", resolution=300.0, save_all=True, append_images=imgs[1:])
    print(f"Wrote {pdf_path}  ({len(imgs)} pages) + page PNGs in {print_dir}")


if __name__ == "__main__":
    main()

"""Build the printable PDF pack for College Command Center™ (12 pages, US Letter).

  1  Applicant Profile            7  Activities & Awards Résumé
  2  College List                 8  Scholarship Log
  3  Application Tracker           9  Net-Price Comparison
  4  Essay & Supplement Tracker  10  Visits & Interviews
  5  Recommendation Tracker      11  Decisions & Compare
  6  Test Scores                 12  Master To-Do & Deadlines

Outputs ../College_Printables.pdf and page PNGs in ../marketing/print/.
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
    d.text((M, 90), "COLLEGE COMMAND CENTER™", font=fs(34), fill=GOLD_HI, anchor="lt")
    d.text((M, 150), title, font=fserif(74), fill=WHITE, anchor="lt")
    if subtitle:
        d.text((M, 258), subtitle, font=fs(30, bold=False), fill=(214, 226, 222), anchor="lt")
    d.line((M, H - 130, W - M, H - 130), fill=LINE, width=2)
    d.text((M, H - 108), "Ella Bennett  ·  Class of 2027", font=fs(24, bold=False), fill=MUTED, anchor="lt")
    d.text((W - M, H - 108), "College Command Center™", font=fs(24, bold=False), fill=ACCENT, anchor="rt")
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
    img, d = page("Applicant Profile", "Your stats at a glance — what shapes your whole list")
    lw = CW / 2 - 30
    left = ["Applicant name", "Class of", "Weighted GPA", "Unweighted GPA", "Class rank", "Intended major"]
    right = ["Best SAT", "Best ACT", "AP / other scores", "Hooks / spikes", "FAFSA submitted?", "Fee-waiver eligible?"]
    for i in range(6):
        field(d, M, TOP + i * 118, lw, left[i])
        field(d, M + CW / 2 + 30, TOP + i * 118, lw, right[i])
    y = TOP + 6 * 118 + 30
    section(d, M, y, CW, "My story — the themes that tie my application together"); y += 96
    for i in range(4):
        d.line((M, y + i * 84 + 30, M + CW, y + i * 84 + 30), fill=LINE, width=2)
    imgs.append(img)


def p02(imgs):
    img, d = page("College List", "A balanced list — reach, match & safety")
    table(d, M, TOP, CW, ["COLLEGE", "R/M/S", "DEADLINE", "APP TYPE", "FEE"],
          [0.0, 0.40, 0.54, 0.72, 0.90], 16, rowh=130,
          filled_rows=[[""] for _ in range(16)])
    imgs.append(img)


def p03(imgs):
    img, d = page("Application Tracker", "Check off every piece for every school")
    table(d, M, TOP, CW, ["COLLEGE", "ESSAYS", "RECS", "FORM", "FEE", "SENT"],
          [0.0, 0.40, 0.52, 0.64, 0.76, 0.88], 16, rowh=130)
    imgs.append(img)


def p04(imgs):
    img, d = page("Essay & Supplement Tracker", "Every prompt, word limit & draft status")
    table(d, M, TOP, CW, ["SCHOOL", "PROMPT", "WORDS", "STATUS"],
          [0.0, 0.28, 0.72, 0.86], 18, rowh=120)
    imgs.append(img)


def p05(imgs):
    img, d = page("Recommendation Tracker", "Who's writing, asked & submitted — chase early")
    table(d, M, TOP, CW, ["RECOMMENDER", "ROLE", "ASKED?", "SUBMITTED?"],
          [0.0, 0.36, 0.58, 0.80], 12, rowh=170)
    imgs.append(img)


def p06(imgs):
    img, d = page("Test Scores", "SAT, ACT, AP & more — the scores you'll send")
    table(d, M, TOP, CW, ["TEST", "DATE", "SCORE", "SENT TO", "DETAIL"],
          [0.0, 0.24, 0.42, 0.58, 0.80], 14, rowh=150)
    imgs.append(img)


def p07(imgs):
    img, d = page("Activities & Awards Résumé", "Your extracurricular story — role & years")
    table(d, M, TOP, CW, ["ACTIVITY / AWARD", "ROLE", "GRADES", "HRS/WK"],
          [0.0, 0.44, 0.64, 0.84], 16, rowh=130)
    imgs.append(img)


def p08(imgs):
    img, d = page("Scholarship Log", "Free money is worth the extra essays")
    table(d, M, TOP, CW, ["SCHOLARSHIP", "AMOUNT", "DEADLINE", "STATUS"],
          [0.0, 0.42, 0.60, 0.80], 18, rowh=120)
    imgs.append(img)


def p09(imgs):
    img, d = page("Net-Price Comparison", "Sticker is a myth — compare true cost after aid")
    table(d, M, TOP, CW, ["COLLEGE", "STICKER / YR", "GRANTS & AID", "NET PRICE"],
          [0.0, 0.40, 0.60, 0.82], 12, rowh=170,
          filled_rows=[[""] for _ in range(12)])
    imgs.append(img)


def p10(imgs):
    img, d = page("Visits & Interviews", "Tours & interviews — impressions while fresh")
    table(d, M, TOP, CW, ["DATE", "SCHOOL", "TYPE", "IMPRESSION"],
          [0.0, 0.14, 0.42, 0.64], 16, rowh=130)
    imgs.append(img)


def p11(imgs):
    img, d = page("Decisions & Compare", "Every result — then choose with a clear head")
    table(d, M, TOP, CW, ["COLLEGE", "DECISION", "NET PRICE", "PROS / CONS"],
          [0.0, 0.32, 0.50, 0.68], 12, rowh=170)
    imgs.append(img)


def p12(imgs):
    img, d = page("Master To-Do & Deadlines", "Everything left this season, in order")
    y = TOP
    for i in range(16):
        col = i % 2
        x = M + col * (CW / 2 + 30)
        checkbox(d, x, y + (i // 2) * 150, "", size=44)
        d.line((x + 70, y + (i // 2) * 150 + 44, x + CW / 2 - 30, y + (i // 2) * 150 + 44), fill=LINE, width=2)
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
    pdf_path = os.path.join(out_dir, "College_Printables.pdf")
    imgs[0].save(pdf_path, "PDF", resolution=300.0, save_all=True, append_images=imgs[1:])
    print(f"Wrote {pdf_path}  ({len(imgs)} pages) + page PNGs in {print_dir}")


if __name__ == "__main__":
    main()

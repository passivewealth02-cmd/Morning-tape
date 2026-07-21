"""Build the printable PDF pack for Elementary Command Center™ (12 pages, US Letter).

  1  Student-at-a-Glance          7  Assignment & Grade Sheet
  2  Weekly Lesson Plan           8  Report Card
  3  Subject Mastery Checklist    9  Attendance & Days
  4  Math Facts Chart            10  Field-Trip & Enrichment Log
  5  Sight Words & Spelling      11  Habits & Character Chart
  6  Reading Log & Levels        12  Awards & Milestones

Outputs ../Elementary_Printables.pdf and page PNGs in ../marketing/print/.
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
    d.text((M, 90), "ELEMENTARY COMMAND CENTER™", font=fs(34), fill=GOLD_HI, anchor="lt")
    d.text((M, 150), title, font=fserif(78), fill=WHITE, anchor="lt")
    if subtitle:
        d.text((M, 258), subtitle, font=fs(30, bold=False), fill=(214, 226, 222), anchor="lt")
    d.line((M, H - 130, W - M, H - 130), fill=LINE, width=2)
    d.text((M, H - 108), "The Bennett Family  ·  2026–2027", font=fs(24, bold=False), fill=MUTED, anchor="lt")
    d.text((W - M, H - 108), "Elementary Command Center™", font=fs(24, bold=False), fill=ACCENT, anchor="rt")
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
    img, d = page("Student-at-a-Glance", "One page per child — grade, level, strengths & goals")
    d.rounded_rectangle((W - M - 420, TOP, W - M, TOP + 520), radius=18, outline=GOLD_LT, width=4)
    d.text((W - M - 210, TOP + 260), "photo", font=fs(30, bold=False), fill=MUTED, anchor="mm")
    lw = CW - 480
    for i, lab in enumerate(["Student name", "Grade & age", "Reading level", "Strengths",
                             "Working on", "Favorite subject"]):
        field(d, M, TOP + i * 92, lw, lab)
    y = TOP + 560
    section(d, M, y, CW, "This year's subjects & curriculum"); y += 96
    for i in range(8):
        col = i % 2
        field(d, M + col * (CW / 2 + 30), y, CW / 2 - 30, "Subject / curriculum")
        if col == 1:
            y += 92
    y += 20
    section(d, M, y, CW, "Goals for the year"); y += 96
    for i in range(3):
        checkbox(d, M, y + i * 84, "", size=40)
        d.line((M + 66, y + i * 84 + 40, M + CW, y + i * 84 + 40), fill=LINE, width=2)
    imgs.append(img)


def p02(imgs):
    img, d = page("Weekly Lesson Plan", "Map the week, subject by subject — then check it off")
    days = ["Mon", "Tue", "Wed", "Thu", "Fri"]
    subs = ["Math", "Reading", "Writing", "Spelling", "Science", "Soc. Studies", "Art", "PE"]
    gx, gy = M, TOP + 10
    gw = CW; gh = 2500
    colw = gw / (len(days) + 1)
    rowh = gh / (len(subs) + 1)
    d.rounded_rectangle((gx, gy, gx + gw, gy + rowh), radius=8, fill=PRIMARY)
    for i, dn in enumerate(days):
        d.text((gx + colw * (i + 1) + colw / 2, gy + rowh / 2), dn, font=fs(28), fill=WHITE, anchor="mm")
    for r, rn in enumerate(subs):
        ry = gy + rowh * (r + 1)
        d.rectangle((gx, ry, gx + colw, ry + rowh), fill=SURFACE)
        d.text((gx + 20, ry + rowh / 2), rn, font=fs(24), fill=PRIMARY, anchor="lm")
    for i in range(len(days) + 2):
        d.line((gx + colw * i, gy, gx + colw * i, gy + gh), fill=LINE, width=2)
    for r in range(len(subs) + 2):
        d.line((gx, gy + rowh * r, gx + gw, gy + rowh * r), fill=LINE, width=2)
    imgs.append(img)


def p03(imgs):
    img, d = page("Subject Mastery Checklist", "The skills that matter — mastered, practicing or not yet")
    table(d, M, TOP, CW, ["SUBJECT", "SKILL", "STUDENT", "M / P / N"],
          [0.0, 0.22, 0.68, 0.86], 20, rowh=130)
    imgs.append(img)


def p04(imgs):
    img, d = page("Math Facts Fluency Chart", "Color a box as each set of facts becomes fast & automatic")
    ops = ["Addition", "Subtraction", "Multiplication", "Division"]
    y = TOP
    field(d, M, y, CW / 2 - 30, "Student"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Date")
    y += 130
    for op in ops:
        section(d, M, y, CW, op); y += 80
        cols = 13
        bw = (CW - (cols - 1) * 18) / cols
        for i in range(cols):
            bx = M + i * (bw + 18)
            d.rounded_rectangle((bx, y, bx + bw, y + bw), radius=8, outline=PRIMARY, width=3)
            d.text((bx + bw / 2, y + bw / 2), str(i), font=fs(28), fill=MUTED, anchor="mm")
        y += bw + 46
    imgs.append(img)


def p05(imgs):
    img, d = page("Sight Words & Spelling", "Every list & unit — known vs total, one row at a time")
    table(d, M, TOP, CW, ["LIST / UNIT", "STUDENT", "KNOWN", "TOTAL", "%"],
          [0.0, 0.44, 0.62, 0.76, 0.90], 18, rowh=120)
    imgs.append(img)


def p06(imgs):
    img, d = page("Reading Log & Levels", "Every book, per child, with its level — reading is the heart of it")
    d.text((M, TOP), "STUDENT:", font=fs(26), fill=ACCENT, anchor="lt")
    d.line((M + 200, TOP + 34, M + 900, TOP + 34), fill=LINE, width=2)
    d.text((M + 1000, TOP), "GOAL:", font=fs(26), fill=ACCENT, anchor="lt")
    d.line((M + 1140, TOP + 34, M + 1500, TOP + 34), fill=LINE, width=2)
    table(d, M, TOP + 90, CW, ["#", "DATE", "TITLE", "LEVEL", "★"],
          [0.0, 0.08, 0.24, 0.80, 0.92], 18, rowh=120,
          filled_rows=[[str(i + 1)] for i in range(18)])
    imgs.append(img)


def p07(imgs):
    img, d = page("Assignment & Grade Sheet", "A clean record of graded work — the source of report cards")
    d.text((M, TOP), "STUDENT:", font=fs(26), fill=ACCENT, anchor="lt")
    d.line((M + 200, TOP + 34, M + 900, TOP + 34), fill=LINE, width=2)
    d.text((M + 1000, TOP), "SUBJECT:", font=fs(26), fill=ACCENT, anchor="lt")
    d.line((M + 1200, TOP + 34, M + 1600, TOP + 34), fill=LINE, width=2)
    table(d, M, TOP + 90, CW, ["DATE", "ASSIGNMENT", "SUBJECT", "GRADE"],
          [0.0, 0.16, 0.60, 0.84], 18, rowh=120)
    imgs.append(img)


def p08(imgs):
    img, d = page("Report Card", "A tidy quarter-by-quarter record — print, sign & file")
    y = TOP
    field(d, M, y, CW / 2 - 30, "Student"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Grade & year")
    y += 130
    table(d, M, y, CW, ["SUBJECT", "Q1", "Q2", "Q3", "Q4", "FINAL"],
          [0.0, 0.40, 0.52, 0.64, 0.76, 0.88], 9, rowh=120,
          filled_rows=[["Math"], ["Reading"], ["Writing"], ["Spelling"], ["Science"],
                       ["Social Studies"], ["Art"], ["PE"]])
    y2 = y + 78 + 9 * 120 + 50
    section(d, M, y2, CW, "Teacher comments"); y2 += 96
    for i in range(2):
        d.line((M, y2 + i * 80 + 30, M + CW, y2 + i * 80 + 30), fill=LINE, width=2)
    imgs.append(img)


def p09(imgs):
    img, d = page("Attendance & Days", "Days by week — the record your state may require")
    table(d, M, TOP, CW, ["WEEK", "M", "T", "W", "TH", "F", "DAYS"],
          [0.0, 0.24, 0.37, 0.50, 0.63, 0.76, 0.88], 12, rowh=170,
          filled_rows=[[f"Week {i+1}"] for i in range(12)])
    imgs.append(img)


def p10(imgs):
    img, d = page("Field-Trip & Enrichment Log", "Learning beyond the table — trips, classes & tie-ins")
    table(d, M, TOP, CW, ["DATE", "PLACE / ACTIVITY", "SUBJECT TIE-IN", "WHO", "★"],
          [0.0, 0.16, 0.52, 0.74, 0.90], 16, rowh=130)
    imgs.append(img)


def p11(imgs):
    img, d = page("Habits & Character Chart", "The little routines that build great learners — check them off")
    habits = ["Made the bed", "Morning chores", "Read 20 minutes", "Kind words / helped",
              "Screen-time limit", "Tidied school space", "Practiced an instrument", "Outside play"]
    days = ["M", "T", "W", "Th", "F", "S", "S"]
    gx = M + 560; gw = CW - 560
    colw = gw / 7
    d.text((M, TOP - 4), "HABIT", font=fs(24), fill=ACCENT, anchor="lt")
    for i, dn in enumerate(days):
        d.text((gx + colw * i + colw / 2, TOP + 8), dn, font=fs(26), fill=PRIMARY, anchor="mm")
    y = TOP + 60
    for h in habits:
        d.rounded_rectangle((M, y, M + 520, y + 76), radius=8, fill=SURFACE)
        d.text((M + 20, y + 38), h, font=fs(24), fill=PRIMARY, anchor="lm")
        for i in range(7):
            bx = gx + colw * i + colw / 2 - 30
            d.rounded_rectangle((bx, y + 8, bx + 60, y + 68), radius=8, outline=PRIMARY, width=3)
        y += 110
    imgs.append(img)


def p12(imgs):
    img, d = page("Awards & Milestones", "Every win worth celebrating — the moments that keep them motivated")
    y = TOP
    for i in range(8):
        d.rounded_rectangle((M, y, M + 110, y + 110), radius=14, outline=GOLD_LT, width=4)
        d.text((M + 55, y + 55), "★", font=fs(50), fill=GOLD_LT, anchor="mm")
        field(d, M + 150, y + 6, CW - 300, "Award / milestone")
        field(d, M + 150, y + 6, 0, "", line=False)
        d.text((W - M, y + 90), "date: __________", font=fs(24, bold=False), fill=MUTED, anchor="rt")
        y += 200
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
    pdf_path = os.path.join(out_dir, "Elementary_Printables.pdf")
    imgs[0].save(pdf_path, "PDF", resolution=300.0, save_all=True, append_images=imgs[1:])
    print(f"Wrote {pdf_path}  ({len(imgs)} pages) + page PNGs in {print_dir}")


if __name__ == "__main__":
    main()

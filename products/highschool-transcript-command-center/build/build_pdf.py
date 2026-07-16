"""Printable PDF pack for High School Transcript Command Center™ (12 pages, US Letter).

The hero page is a real, print-ready OFFICIAL TRANSCRIPT built from the same course
data as the workbook, so the numbers always match.

  1  Official Transcript          7  Awards & Honors
  2  4-Year Plan                  8  Community Service Log
  3  Course Records / Grade Sheet 9  Course Descriptions
  4  GPA Worksheet               10  Reading List
  5  Credit & Grad Checklist     11  College Application Tracker
  6  Test Score Record           12  Activities & Leadership Résumé

Outputs ../HS_Transcript_Printables.pdf and page PNGs in ../marketing/print/.
Run: python3 build_pdf.py
"""
from __future__ import annotations
import os
from PIL import Image, ImageDraw, ImageFont
from build_xlsx import COURSES, GRADE_SCALE

PRIMARY = (27, 79, 72); PRIMARY_DK = (18, 56, 51); ACCENT = (147, 115, 86)
GOLD = (180, 145, 90); GOLD_LT = (201, 168, 106); GOLD_HI = (224, 196, 140)
SURFACE = (229, 211, 186); MINT = (227, 248, 239); WARN = (251, 240, 226)
WHITE = (255, 255, 255); TEXT = (51, 51, 51); MUTED = (120, 114, 104)
LINE = (200, 194, 182); ROW_ALT = (247, 243, 236)

W, H = 2550, 3300
M = 190
GP = dict(GRADE_SCALE)
BUMP = {"Regular": 0.0, "Honors": 0.5, "AP": 1.0, "Dual Credit": 1.0}

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
    d.text((M, 90), "HIGH SCHOOL TRANSCRIPT COMMAND CENTER™", font=fs(32), fill=GOLD_HI, anchor="lt")
    d.text((M, 150), title, font=fserif(74), fill=WHITE, anchor="lt")
    if subtitle:
        d.text((M, 256), subtitle, font=fs(30, bold=False), fill=(214, 226, 222), anchor="lt")
    d.line((M, H - 130, W - M, H - 130), fill=LINE, width=2)
    d.text((M, H - 108), "Ella Bennett  ·  Class of 2027", font=fs(24, bold=False), fill=MUTED, anchor="lt")
    d.text((W - M, H - 108), "HS Transcript Command Center™", font=fs(24, bold=False), fill=ACCENT, anchor="rt")
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


def _gpa():
    cr = sum(c[4] for c in COURSES)
    uw = sum(GP[c[5]] * c[4] for c in COURSES) / cr
    w = sum((GP[c[5]] + BUMP[c[3]]) * c[4] for c in COURSES) / cr
    return cr, uw, w


# ---------------------------------------------------------------- pages
def p01(imgs):
    # The hero: an official-looking transcript
    img = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(img)
    d.rectangle((M - 40, 150, W - M + 40, H - 150), outline=GOLD_LT, width=4)
    d.text((W // 2, 240), "Bennett Family Academy", font=fserif(60), fill=PRIMARY, anchor="mm")
    d.text((W // 2, 320), "OFFICIAL HIGH SCHOOL TRANSCRIPT", font=fs(30), fill=ACCENT, anchor="mm")
    d.line((M, 380, W - M, 380), fill=GOLD_LT, width=3)
    # student info
    info = [("Student", "Ella Bennett"), ("Class Of", "2027"), ("Graduation", "May 28, 2027"),
            ("Grading Scale", "4.0 · Weighted (H +0.5, AP +1.0)")]
    ix = M + 20
    for i, (a, b) in enumerate(info):
        yy = 430 + (i // 2) * 70
        xx = M + 20 + (i % 2) * (CW // 2)
        d.text((xx, yy), a.upper() + ":", font=fs(22), fill=ACCENT, anchor="lt")
        d.text((xx + 230, yy), b, font=fs(24, bold=False), fill=TEXT, anchor="lt")
    # two columns of year blocks
    colw = (CW - 60) / 2
    y0 = 610
    for ci, years in enumerate([(9, 10), (11, 12)]):
        cx = M + 20 + ci * (colw + 60)
        yy = y0
        for yr in years:
            d.rounded_rectangle((cx, yy, cx + colw, yy + 44), radius=6, fill=SURFACE)
            d.text((cx + 16, yy + 22), f"GRADE {yr}", font=fs(24), fill=PRIMARY, anchor="lm")
            yy += 58
            d.text((cx + 16, yy), "COURSE", font=fs(18), fill=ACCENT, anchor="lt")
            d.text((cx + colw - 150, yy), "CR", font=fs(18), fill=ACCENT, anchor="lt")
            d.text((cx + colw - 70, yy), "GR", font=fs(18), fill=ACCENT, anchor="lt")
            yy += 34
            d.line((cx, yy, cx + colw, yy), fill=LINE, width=2); yy += 8
            for (y, area, course, level, cr, grade) in [c for c in COURSES if c[0] == yr]:
                title = course + (" (H)" if level == "Honors" else " (AP)" if level == "AP" else "")
                d.text((cx + 16, yy), title, font=fs(21, bold=False), fill=TEXT, anchor="lt")
                d.text((cx + colw - 150, yy), f"{cr:.1f}", font=fs(21, bold=False), fill=TEXT, anchor="lt")
                d.text((cx + colw - 70, yy), grade, font=fs(21), fill=PRIMARY, anchor="lt")
                yy += 40
            yc = sum(c[4] for c in COURSES if c[0] == yr)
            d.line((cx, yy, cx + colw, yy), fill=LINE, width=1); yy += 8
            d.text((cx + 16, yy), f"Year credits: {yc:.1f}", font=fs(20), fill=ACCENT, anchor="lt")
            yy += 66
    # summary box
    cr, uw, w = _gpa()
    by = 2560
    d.rounded_rectangle((M, by, W - M, by + 300), radius=12, fill=MINT, outline=GOLD_LT, width=2)
    d.text((M + 40, by + 30), "SUMMARY", font=fs(28), fill=PRIMARY, anchor="lt")
    stats = [("Cumulative GPA (Unweighted)", f"{uw:.2f}"), ("Cumulative GPA (Weighted)", f"{w:.2f}"),
             ("Total Credits", f"{cr:.1f}"), ("Class Of", "2027")]
    for i, (a, b) in enumerate(stats):
        xx = M + 40 + (i % 2) * (CW // 2)
        yy = by + 100 + (i // 2) * 80
        d.text((xx, yy), a, font=fs(22, bold=False), fill=TEXT, anchor="lt")
        d.text((xx + 520, yy), b, font=fs(28), fill=PRIMARY, anchor="lt")
    d.text((M + 40, by + 250), "Administrator signature: ______________________________    Date: ____________",
           font=fs(22, bold=False), fill=TEXT, anchor="lt")
    imgs.append(img)


def p02(imgs):
    img, d = page("4-Year Plan", "Map all four years — spot credit & requirement gaps early")
    table(d, M, TOP, CW, ["SUBJECT", "GRADE 9", "GRADE 10", "GRADE 11", "GRADE 12"],
          [0.0, 0.22, 0.42, 0.62, 0.82], 9, rowh=200,
          filled_rows=[["English"], ["Math"], ["Science"], ["Social Studies"], ["World Language"], ["Electives"]])
    imgs.append(img)


def p03(imgs):
    img, d = page("Course Records / Grade Sheet", "Every course, credit & grade — the source of the GPA")
    table(d, M, TOP, CW, ["YR", "SUBJECT", "COURSE", "LEVEL", "CR", "GRADE"],
          [0.0, 0.10, 0.34, 0.66, 0.80, 0.90], 20, rowh=118)
    imgs.append(img)


def p04(imgs):
    img, d = page("GPA Worksheet", "Grade points × credits — unweighted & weighted, by year")
    d.text((M, TOP), "SCALE: A=4.0  A-=3.7  B+=3.3  B=3.0  B-=2.7  C+=2.3  C=2.0   ·   Honors +0.5  ·  AP +1.0",
           font=fs(24, bold=False), fill=TEXT, anchor="lt")
    table(d, M, TOP + 70, CW, ["COURSE", "CREDITS", "GRADE", "GPA PTS", "QUALITY PTS"],
          [0.0, 0.46, 0.60, 0.74, 0.86], 16, rowh=130)
    y = TOP + 70 + 78 + 16 * 130 + 50
    section(d, M, y, CW, "Totals"); y += 96
    for lab in ["Total credits", "Total quality points", "Unweighted GPA", "Weighted GPA"]:
        field(d, M, y, CW / 2 - 30, lab)
        if lab in ("Total quality points", "Weighted GPA"):
            y += 92
        else:
            field(d, M + CW / 2 + 30, y, CW / 2 - 30, {"Total credits": "Cumulative", "Unweighted GPA": "Notes"}.get(lab, ""))
            y += 92
    imgs.append(img)


def p05(imgs):
    img, d = page("Credit & Graduation Checklist", "Credits by area vs requirements — confirm your own state / college")
    table(d, M, TOP, CW, ["SUBJECT AREA", "REQUIRED", "EARNED", "REMAINING"],
          [0.0, 0.46, 0.64, 0.82], 9, rowh=150,
          filled_rows=[["English"], ["Math"], ["Science"], ["Social Studies"], ["World Language"],
                       ["Fine Arts"], ["PE/Health"], ["Electives"], ["TOTAL"]])
    y = TOP + 78 + 9 * 150 + 60
    for it in ["Total 24 credits earned", "Course descriptions written", "Transcript signed & dated",
               "Diploma ordered / issued"]:
        checkbox(d, M, y, it, size=46, font=fs(28, bold=False)); y += 96
    imgs.append(img)


def p06(imgs):
    img, d = page("Test Score Record", "SAT, ACT, PSAT, AP & CLEP — the scores colleges ask for")
    table(d, M, TOP, CW, ["TEST", "DATE", "SCORE", "DETAIL / PERCENTILE"],
          [0.0, 0.30, 0.50, 0.68], 16, rowh=130)
    imgs.append(img)


def p07(imgs):
    img, d = page("Awards & Honors", "Every recognition, dated — ready to drop into applications")
    table(d, M, TOP, CW, ["AWARD / HONOR", "GRADE", "AWARDED BY", "DATE"],
          [0.0, 0.48, 0.62, 0.82], 18, rowh=120)
    imgs.append(img)


def p08(imgs):
    img, d = page("Community Service Log", "Log every hour — scholarships & honor societies require a total")
    table(d, M, TOP, CW, ["DATE", "ORGANIZATION / ACTIVITY", "HOURS", "SUPERVISOR"],
          [0.0, 0.18, 0.64, 0.80], 18, rowh=120)
    y = TOP + 78 + 18 * 120 + 40
    field(d, M, y, CW / 2 - 30, "Total hours"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Goal")
    imgs.append(img)


def p09(imgs):
    img, d = page("Course Descriptions", "Homeschool must-have — a short paragraph per course for admissions")
    y = TOP
    for i in range(6):
        d.rounded_rectangle((M, y, M + CW, y + 44), radius=6, fill=SURFACE)
        d.text((M + 16, y + 22), f"COURSE {i+1}", font=fs(24), fill=PRIMARY, anchor="lm")
        d.line((M + 220, y + 22, M + CW - 16, y + 22), fill=LINE, width=2)
        y += 64
        for k in range(3):
            d.line((M, y + k * 60 + 40, M + CW, y + k * 60 + 40), fill=LINE, width=2)
        y += 3 * 60 + 40
    imgs.append(img)


def p10(imgs):
    img, d = page("Reading List", "Great books read — strengthens English credits & the narrative")
    table(d, M, TOP, CW, ["#", "TITLE", "AUTHOR", "COURSE / YEAR", "DONE"],
          [0.0, 0.08, 0.44, 0.70, 0.90], 20, rowh=130,
          filled_rows=[[str(i + 1)] for i in range(20)])
    imgs.append(img)


def p11(imgs):
    img, d = page("College Application Tracker", "Every school, deadline & piece — nothing missed")
    table(d, M, TOP, CW, ["COLLEGE", "DEADLINE", "APPLIED", "TRANSCRIPT", "ESSAY", "$"],
          [0.0, 0.34, 0.52, 0.66, 0.80, 0.90], 16, rowh=130)
    imgs.append(img)


def p12(imgs):
    img, d = page("Activities & Leadership Résumé", "The extracurricular story — activity, role & years")
    table(d, M, TOP, CW, ["ACTIVITY", "GRADES", "ROLE / DETAIL", "YEARS"],
          [0.0, 0.46, 0.60, 0.88], 16, rowh=130)
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
    pdf_path = os.path.join(out_dir, "HS_Transcript_Printables.pdf")
    imgs[0].save(pdf_path, "PDF", resolution=300.0, save_all=True, append_images=imgs[1:])
    print(f"Wrote {pdf_path}  ({len(imgs)} pages) + page PNGs in {print_dir}")


if __name__ == "__main__":
    main()

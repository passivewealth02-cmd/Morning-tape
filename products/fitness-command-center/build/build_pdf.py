"""Build the printable PDF pack for Fitness & Meal-Prep Command Center™ (12 pages).

  1  Weekly Meal Planner          7  Body Measurements
  2  Grocery List (by aisle)      8  Weight Progress Chart
  3  Macro Tracker                9  Habit Tracker (monthly)
  4  Recipe Cards                10  Meal-Prep Day Checklist
  5  Workout Plan (weekly split) 11  Progress & Measurements
  6  Workout Log                 12  Goals & Why

Outputs ../Fitness_Printables.pdf and page PNGs in ../marketing/print/.
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
    d.text((M, 90), "FITNESS & MEAL-PREP COMMAND CENTER™", font=fs(32), fill=GOLD_HI, anchor="lt")
    d.text((M, 150), title, font=fserif(74), fill=WHITE, anchor="lt")
    if subtitle:
        d.text((M, 258), subtitle, font=fs(30, bold=False), fill=(214, 226, 222), anchor="lt")
    d.line((M, H - 130, W - M, H - 130), fill=LINE, width=2)
    d.text((M, H - 108), "Week of: ______________", font=fs(24, bold=False), fill=MUTED, anchor="lt")
    d.text((W - M, H - 108), "Fitness & Meal-Prep Command Center™", font=fs(24, bold=False), fill=ACCENT, anchor="rt")
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
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]


# ---------------------------------------------------------------- pages
def p01(imgs):
    img, d = page("Weekly Meal Planner", "A week of meals — breakfast, lunch, dinner & snacks")
    table(d, M, TOP, CW, ["DAY", "BREAKFAST", "LUNCH", "DINNER", "SNACK"],
          [0.0, 0.16, 0.40, 0.62, 0.84], 7, rowh=300,
          filled_rows=[[dname] for dname in DAYS])
    imgs.append(img)


def p02(imgs):
    img, d = page("Grocery List", "Everything the plan needs — check it off as you shop")
    half = CW / 2 - 30
    cats = ["Produce", "Meat & Fish", "Dairy", "Pantry", "Frozen", "Other"]
    y = TOP
    for i, cat in enumerate(cats):
        col = i % 2
        x = M + col * (half + 60)
        yy = TOP + (i // 2) * 780
        section(d, x, yy, half, cat)
        for k in range(7):
            checkbox(d, x, yy + 90 + k * 96, "", size=44)
            d.line((x + 68, yy + 90 + k * 96 + 44, x + half, yy + 90 + k * 96 + 44), fill=LINE, width=2)
    imgs.append(img)


def p03(imgs):
    img, d = page("Macro Tracker", "Log calories & protein daily — stay near your target")
    table(d, M, TOP, CW, ["DATE", "CALORIES", "PROTEIN", "CARBS", "FAT", "NOTES"],
          [0.0, 0.14, 0.30, 0.46, 0.60, 0.74], 20, rowh=118)
    imgs.append(img)


def p04(imgs):
    img, d = page("Recipe Cards", "Your go-to meals with calories & protein per serving")
    half = CW / 2 - 30
    for i in range(4):
        col = i % 2
        x = M + col * (half + 60)
        y = TOP + (i // 2) * 1220
        d.rounded_rectangle((x, y, x + half, y + 1120), radius=16, outline=PRIMARY, width=3)
        d.rounded_rectangle((x, y, x + half, y + 90), radius=16, fill=SURFACE)
        d.text((x + 24, y + 45), f"RECIPE {i+1}", font=fs(28), fill=PRIMARY, anchor="lm")
        field(d, x + 24, y + 130, half - 48, "Name")
        field(d, x + 24, y + 250, (half - 72) / 2, "Servings")
        field(d, x + 24 + (half - 48) / 2, y + 250, (half - 72) / 2, "Cal / serving")
        field(d, x + 24, y + 370, (half - 72) / 2, "Protein (g)")
        field(d, x + 24 + (half - 48) / 2, y + 370, (half - 72) / 2, "Category")
        d.text((x + 24, y + 470), "INGREDIENTS / STEPS", font=fs(24), fill=ACCENT, anchor="lt")
        for k in range(6):
            d.line((x + 24, y + 560 + k * 88, x + half - 24, y + 560 + k * 88), fill=LINE, width=2)
    imgs.append(img)


def p05(imgs):
    img, d = page("Workout Plan", "Your weekly split — check off each session")
    table(d, M, TOP, CW, ["DAY", "FOCUS", "EXERCISES", "DONE?"],
          [0.0, 0.16, 0.40, 0.86], 7, rowh=280,
          filled_rows=[[dname] for dname in DAYS])
    imgs.append(img)


def p06(imgs):
    img, d = page("Workout Log", "Every lift — sets × reps × weight = volume")
    table(d, M, TOP, CW, ["DATE", "EXERCISE", "SETS", "REPS", "WEIGHT", "VOLUME"],
          [0.0, 0.14, 0.48, 0.60, 0.72, 0.85], 20, rowh=118)
    imgs.append(img)


def p07(imgs):
    img, d = page("Body Measurements", "Track more than the scale — inches tell the real story")
    labels = ["Weight (lb)", "Chest", "Waist", "Hips", "Left arm", "Right arm",
              "Left thigh", "Right thigh", "Body fat %"]
    table(d, M, TOP, CW, ["MEASUREMENT", "START", "WEEK 4", "WEEK 8", "WEEK 12"],
          [0.0, 0.40, 0.55, 0.70, 0.85], len(labels), rowh=200,
          filled_rows=[[la] for la in labels])
    imgs.append(img)


def p08(imgs):
    img, d = page("Weight Progress Chart", "Plot your weekly weigh-in — watch the trend")
    gx0, gy0 = M + 140, TOP + 40
    gx1, gy1 = W - M, H - 460
    d.line((gx0, gy0, gx0, gy1), fill=PRIMARY, width=4)
    d.line((gx0, gy1, gx1, gy1), fill=PRIMARY, width=4)
    for i in range(9):
        yy = gy0 + (gy1 - gy0) * i / 8
        d.line((gx0 - 20, yy, gx1, yy), fill=LINE, width=1)
        d.text((gx0 - 40, yy), "____", font=fs(22, bold=False), fill=MUTED, anchor="rm")
    d.text((gx0 - 120, (gy0 + gy1) / 2), "WEIGHT", font=fs(24), fill=ACCENT, anchor="mm")
    for i in range(12):
        xx = gx0 + (gx1 - gx0) * (i + 0.5) / 12
        d.text((xx, gy1 + 30), f"W{i+1}", font=fs(24, bold=False), fill=MUTED, anchor="mm")
    imgs.append(img)


def p09(imgs):
    img, d = page("Habit Tracker", "Water, sleep, steps & workouts — color a box a day")
    habits = ["Water (8 cups)", "Sleep (7+ hrs)", "10k steps", "Workout", "Hit protein", "No added sugar"]
    gx = M + 360
    bw = (CW - 360) / 31
    for dn in range(31):
        d.text((gx + bw * dn + bw / 2, TOP + 10), str(dn + 1), font=fs(18, bold=False), fill=MUTED, anchor="mm")
    for i, hb in enumerate(habits):
        y = TOP + 70 + i * 130
        d.text((M, y + 40), hb, font=fs(28, bold=False), fill=PRIMARY, anchor="lm")
        for dn in range(31):
            x = gx + bw * dn
            d.rounded_rectangle((x + 4, y, x + bw - 4, y + 80), radius=8, outline=PRIMARY, width=2)
    imgs.append(img)


def p10(imgs):
    img, d = page("Meal-Prep Day", "Prep once, eat all week — the Sunday checklist")
    y = TOP
    section(d, M, y, CW, "This week's prep"); y += 100
    for task in ["Cook protein (chicken / beef / fish)", "Cook grains (rice / quinoa / oats)",
                 "Roast / chop vegetables", "Portion into containers", "Make snacks & smoothie packs",
                 "Wash & prep fruit", "Label with day & macros"]:
        checkbox(d, M, y, task, size=50, font=fs(34, bold=False))
        y += 120
    y += 20
    field(d, M, y, CW / 2 - 30, "Containers needed")
    field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Prep time")
    imgs.append(img)


def p11(imgs):
    img, d = page("Progress & Measurements", "Photos & numbers — proof the work is working")
    half = CW / 2 - 30
    for i, lab in enumerate(["FRONT", "SIDE"]):
        x = M + i * (half + 60)
        d.rounded_rectangle((x, TOP, x + half, TOP + 1000), radius=16, outline=LINE, width=3)
        d.text((x + half / 2, TOP + 500), lab, font=fs(40), fill=(210, 204, 190), anchor="mm")
    y = TOP + 1080
    for lab in ["Date", "Weight", "Waist", "How I feel (energy, sleep, strength)"]:
        field(d, M, y, CW, lab); y += 150
    imgs.append(img)


def p12(imgs):
    img, d = page("Goals & Why", "The reason behind the reps — read it on the hard days")
    y = TOP
    for who in ["My goal (be specific)", "By when", "Why it matters to me", "How I'll celebrate"]:
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
    pdf_path = os.path.join(out_dir, "Fitness_Printables.pdf")
    imgs[0].save(pdf_path, "PDF", resolution=300.0, save_all=True, append_images=imgs[1:])
    print(f"Wrote {pdf_path}  ({len(imgs)} pages) + page PNGs in {print_dir}")


if __name__ == "__main__":
    main()

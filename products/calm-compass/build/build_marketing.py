"""Premium marketing image set for Calm Compass™ (6 images, 2000x2000).

  01_hero.png       - feature-forward main thumbnail
  02_dashboard.png  - wellness dashboard close-up
  03_checkin.png    - daily check-in
  04_habits.png     - habit tracker + streaks
  05_progress.png   - wellness score + snapshot
  06_mobile.png     - mobile preview

Run: python3 build_marketing.py
"""
from __future__ import annotations
import math
import os
from PIL import Image, ImageDraw, ImageFilter, ImageFont

PRIMARY = (27, 79, 72)
ACCENT = (147, 115, 86)
GOLD = (180, 145, 90)
GOLD_LT = (201, 168, 106)
SURFACE = (229, 211, 186)
HIGHLIGHT = (117, 230, 193)
BG = (251, 248, 242)
WHITE = (255, 255, 255)
TEXT = (51, 51, 51)
TEXT_MUTED = (130, 125, 115)
DANGER = (201, 76, 76)
MINT_BG = (227, 248, 239)
WARN_BG = (251, 240, 226)
RED_BG = (251, 230, 230)
DOT = (224, 216, 202)
SIZE = 2000

SANS_B = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
SANS_R = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
SERIF_B = "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf"


def fs(s, bold=True):
    return ImageFont.truetype(SANS_B if bold else SANS_R, s)


def fserif(s):
    return ImageFont.truetype(SERIF_B, s)


def dotted_bg(c, spacing=46, r=4):
    d = ImageDraw.Draw(c)
    for y in range(spacing // 2, c.height, spacing):
        for x in range(spacing // 2, c.width, spacing):
            d.ellipse((x - r, y - r, x + r, y + r), fill=DOT)


def shadow(c, box, radius, blur=24, alpha=70):
    layer = Image.new("RGBA", c.size, (0, 0, 0, 0))
    ImageDraw.Draw(layer).rounded_rectangle(box, radius=radius, fill=(27, 79, 72, alpha))
    c.alpha_composite(layer.filter(ImageFilter.GaussianBlur(blur)), (0, 18))


def tsize(d, t, f):
    b = d.textbbox((0, 0), t, font=f)
    return b[2] - b[0], b[3] - b[1]


def tc(d, xy, t, f, fill, anchor="mm"):
    d.text(xy, t, font=f, fill=fill, anchor=anchor)


def pill(c, cx, cy, text, font, pad_x=60, pad_y=26, star=False, bg=PRIMARY, fg=WHITE):
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    label = f"★  {text}" if star else text
    tw, th = tsize(d, label, font)
    w, h = tw + pad_x * 2, th + pad_y * 2
    box = (cx - w // 2, cy - h // 2, cx + w // 2, cy + h // 2)
    d.rounded_rectangle(box, radius=h // 2, fill=bg)
    d.rounded_rectangle((box[0] + 4, box[1] + 4, box[2] - 4, box[3] - 4),
                        radius=(h - 8) // 2, outline=GOLD_LT, width=2)
    d.text((cx, cy), label, font=font, fill=fg, anchor="mm")
    c.alpha_composite(ov)


def gold_divider(c, cx, cy, width=560):
    d = ImageDraw.Draw(c)
    d.line((cx - width // 2, cy, cx - 30, cy), fill=GOLD, width=3)
    d.line((cx + 30, cy, cx + width // 2, cy), fill=GOLD, width=3)
    d.polygon([(cx, cy - 12), (cx + 16, cy), (cx, cy + 12), (cx - 16, cy)], fill=GOLD)


def compass_crest(c, cx, cy, r=58):
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    d.ellipse((cx - r, cy - r, cx + r, cy + r), fill=PRIMARY, outline=GOLD_LT, width=4)
    d.ellipse((cx - r + 10, cy - r + 10, cx + r - 10, cy + r - 10), outline=GOLD_LT, width=2)
    # 4-point compass needle
    L = r * 0.55
    s = r * 0.18
    d.polygon([(cx, cy - L), (cx + s, cy), (cx, cy + L), (cx - s, cy)], fill=HIGHLIGHT)
    d.polygon([(cx - L, cy), (cx, cy - s), (cx + L, cy), (cx, cy + s)], fill=GOLD_LT)
    d.ellipse((cx - 7, cy - 7, cx + 7, cy + 7), fill=WHITE)
    c.alpha_composite(ov)


def benefit_badge(c, cx, cy, big, small, w=440, h=170, big_color=GOLD_LT):
    box = (cx - w // 2, cy - h // 2, cx + w // 2, cy + h // 2)
    shadow(c, box, 22, 26, 90)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    d.rounded_rectangle(box, radius=22, fill=PRIMARY, outline=GOLD_LT, width=3)
    d.text((cx, cy - h * 0.20), big, font=fserif(58), fill=big_color, anchor="mm")
    d.text((cx, cy + h * 0.26), small, font=fs(24), fill=WHITE, anchor="mm")
    c.alpha_composite(ov)


def feature_card(c, x, y, w, h, title, sub):
    box = (x, y, x + w, y + h)
    shadow(c, box, 16, 18, 60)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    d.rounded_rectangle(box, radius=16, fill=WHITE, outline=GOLD_LT, width=2)
    cyc = y + h // 2
    bx = x + 44
    d.ellipse((bx - 30, cyc - 30, bx + 30, cyc + 30), fill=HIGHLIGHT, outline=PRIMARY, width=3)
    d.text((bx, cyc - 1), "✓", font=fs(34), fill=PRIMARY, anchor="mm")
    d.text((x + 92, cyc - 22), title, font=fs(29), fill=PRIMARY, anchor="lm")
    d.text((x + 92, cyc + 24), sub, font=fs(22, bold=False), fill=TEXT_MUTED, anchor="lm")
    c.alpha_composite(ov)


def page_scaffold(img, badge, title, subtitle):
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 175, badge, font=fs(42), pad_x=58, pad_y=24)
    tc(d, (SIZE // 2, 310), title, fserif(72), PRIMARY)
    gold_divider(img, SIZE // 2, 380, width=520)
    tc(d, (SIZE // 2, 430), subtitle, fs(31, bold=False), TEXT_MUTED)


def panel(img, box, fill_color=WHITE, title=None):
    shadow(img, box, 24, 38, 105)
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    od.rounded_rectangle(box, radius=24, fill=fill_color, outline=GOLD_LT, width=3)
    img.alpha_composite(ov)
    if title:
        ImageDraw.Draw(img).text((box[0] + 34, box[1] + 26), title, font=fs(30), fill=ACCENT, anchor="lt")


def dashboard_screen(c, screen, charts_note=True):
    sx0, sy0, sx1, sy1 = screen
    sw = sx1 - sx0
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    head = 84
    d.rounded_rectangle((sx0, sy0, sx1, sy0 + head), radius=8, fill=PRIMARY)
    d.rectangle((sx0, sy0 + 30, sx1, sy0 + head), fill=PRIMARY)
    d.text((sx0 + 28, sy0 + head // 2), "CALM COMPASS", font=fs(26), fill=WHITE, anchor="lm")
    d.text((sx1 - 28, sy0 + head // 2), "Your wellness dashboard", font=fs(22), fill=GOLD_LT, anchor="rm")
    d.rectangle((sx0, sy0 + head, sx1, sy0 + head + 4), fill=GOLD_LT)

    gx, gy, gap = sx0 + 22, sy0 + head + 22, 14
    kw = (sw - 22 * 2 - gap * 3) // 4
    kh = 120
    kpis = [
        ("AVG MOOD", "6.8", PRIMARY),
        ("ROUTINE SCORE", "71%", PRIMARY),
        ("AVG SLEEP", "7.3 hrs", PRIMARY),
        ("JOURNAL STREAK", "9", ACCENT),
        ("EXERCISE", "11", PRIMARY),
        ("MINDFUL MIN", "145", PRIMARY),
        ("GOALS DONE", "1", PRIMARY),
        ("CHECK-INS", "14", PRIMARY),
    ]
    for i, (lab, val, col) in enumerate(kpis):
        r, ci = divmod(i, 4)
        x0 = gx + ci * (kw + gap)
        y0 = gy + r * (kh + gap)
        d.rounded_rectangle((x0, y0, x0 + kw, y0 + kh), radius=10, fill=WHITE, outline=(225, 218, 205), width=2)
        d.text((x0 + 14, y0 + 16), lab, font=fs(15), fill=ACCENT, anchor="lt")
        d.text((x0 + 14, y0 + 70), val, font=fserif(36), fill=col, anchor="lm")

    nav_y = gy + 2 * (kh + gap) + 8
    nav_h = 34
    nav = ["Check-In", "Habits", "Planner", "Reflection", "Goals", "Progress"]
    cw = (sw - 22 * 2 - 6 * (len(nav) - 1)) / len(nav)
    for i, name in enumerate(nav):
        x0 = gx + i * (cw + 6)
        d.rounded_rectangle((x0, nav_y, x0 + cw, nav_y + nav_h), radius=17, fill=PRIMARY)
        d.text((x0 + cw / 2, nav_y + nav_h / 2), name, font=fs(15), fill=WHITE, anchor="mm")

    cy0 = nav_y + nav_h + 22
    if charts_note and (sy1 - cy0) > 60:
        # a simple mood trend line inside the screen
        cbox = (gx, cy0, sx1 - 22, sy1 - 18)
        d.rounded_rectangle(cbox, radius=10, fill=WHITE, outline=(225, 218, 205), width=2)
        d.text((cbox[0] + 16, cbox[1] + 14), "MOOD & ENERGY TREND", font=fs(18), fill=ACCENT, anchor="lt")
        lx0, ly0 = cbox[0] + 40, cbox[1] + 60
        lx1, ly1 = cbox[2] - 30, cbox[3] - 30
        import random
        rnd = random.Random(5)
        for series_col, base in [(PRIMARY, 0.55), (HIGHLIGHT, 0.4)]:
            pts = []
            n = 14
            v = base
            for i in range(n):
                v = max(0.1, min(0.95, v + rnd.uniform(-0.12, 0.16)))
                pts.append((lx0 + i / (n - 1) * (lx1 - lx0), ly1 - v * (ly1 - ly0)))
            d.line(pts, fill=series_col, width=5)
            for p in pts[::2]:
                d.ellipse((p[0] - 5, p[1] - 5, p[0] + 5, p[1] + 5), fill=series_col)
    else:
        d.text((sx0 + sw / 2, cy0 + 6),
               "Check-In · Habits · Sleep · Reflection · Gratitude · Progress — all in one file",
               font=fs(18, bold=False), fill=TEXT_MUTED, anchor="mt")
    c.alpha_composite(ov)


def laptop(c, cx, cy, w, h, charts_note=True):
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    sbox = (cx - w // 2, cy - h // 2 + 30, cx + w // 2, cy + h // 2 + 80)
    sh = Image.new("RGBA", c.size, (0, 0, 0, 0))
    ImageDraw.Draw(sh).rounded_rectangle(sbox, radius=40, fill=(27, 79, 72, 110))
    c.alpha_composite(sh.filter(ImageFilter.GaussianBlur(50)))
    body = (cx - w // 2, cy - h // 2, cx + w // 2, cy + h // 2)
    d.rounded_rectangle(body, radius=26, fill=(38, 38, 42), outline=(20, 20, 25), width=4)
    bez = 22
    screen = (body[0] + bez, body[1] + bez, body[2] - bez, body[3] - bez)
    d.rounded_rectangle(screen, radius=8, fill=BG)
    d.rounded_rectangle((cx - w // 2 - 120, cy + h // 2, cx + w // 2 + 120, cy + h // 2 + 24), radius=12, fill=(58, 58, 62))
    d.rounded_rectangle((cx - 80, cy + h // 2 + 4, cx + 80, cy + h // 2 + 14), radius=6, fill=(33, 33, 38))
    c.alpha_composite(ov)
    dashboard_screen(c, screen, charts_note=charts_note)


def draw_table(img, inner, headers, rows, status_map=None, col_aligns=None,
               header_font=None, row_font=None):
    layer = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(layer)
    iw = inner[2] - inner[0]
    ih = inner[3] - inner[1]
    cols = len(headers)
    col_w = iw / cols
    hdr_h = 70
    hf = header_font or fs(22)
    rf = row_font or fs(22)
    od.rounded_rectangle((inner[0], inner[1], inner[2], inner[1] + hdr_h), radius=12, fill=PRIMARY)
    od.rectangle((inner[0], inner[1] + 24, inner[2], inner[1] + hdr_h), fill=PRIMARY)
    for i, h in enumerate(headers):
        cx = inner[0] + col_w * (i + 0.5)
        od.text((cx, inner[1] + hdr_h // 2), h, font=hf, fill=WHITE, anchor="mm")
    row_h = (ih - hdr_h - 10) / len(rows)
    for ri, row in enumerate(rows):
        *cells, bg = row
        y0 = inner[1] + hdr_h + ri * row_h
        if bg is None and ri % 2 == 1:
            bg = (244, 236, 222)
        if bg:
            od.rectangle((inner[0] + 3, y0, inner[2] - 3, y0 + row_h), fill=bg)
        for ci, val in enumerate(cells):
            cx = inner[0] + col_w * (ci + 0.5)
            align = (col_aligns or {}).get(ci, "mm")
            color = PRIMARY if ci == 0 else TEXT
            if align == "lm":
                od.text((inner[0] + col_w * ci + 18, y0 + row_h / 2), str(val), font=rf, fill=color, anchor="lm")
            else:
                od.text((cx, y0 + row_h / 2), str(val), font=rf, fill=color, anchor="mm")
    img.alpha_composite(layer)


# ===========================================================================
# 01 HERO
# ===========================================================================
def render_hero(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)
    d = ImageDraw.Draw(img)

    compass_crest(img, SIZE // 2, 150, r=58)
    pill(img, SIZE // 2, 285, "ANXIETY & SELF-CARE WELLNESS PLANNER", font=fs(35), pad_x=54, pad_y=22)

    tc(d, (SIZE // 2, 432), "CALM COMPASS", fserif(148), PRIMARY)
    gold_divider(img, SIZE // 2, 540, width=560)
    tc(d, (SIZE // 2, 600), "Build calm routines · track your mood · celebrate progress",
       fs(30, bold=False), TEXT_MUTED)

    laptop(img, SIZE // 2, 1000, w=1360, h=560, charts_note=False)
    benefit_badge(img, 300, 730, "78%", "WELLNESS SCORE")
    benefit_badge(img, SIZE - 300, 730, "2-in-1", "EXCEL + GOOGLE SHEETS")

    tc(d, (SIZE // 2, 1360), "A CALM SPACE FOR EVERYTHING", fs(34), ACCENT)
    features = [
        ("Daily Check-In", "mood · energy · sleep"),
        ("Habit Tracker", "auto streaks"),
        ("Mood Trends", "see your patterns"),
        ("Sleep Tracker", "rest & quality"),
        ("Reflection Journal", "gentle prompts"),
        ("Social Prep", "confidence planner"),
        ("Wellness Score", "your progress"),
        ("Gratitude", "daily good things"),
    ]
    cols = 4
    margin = 120
    gx, gy = 24, 24
    cw = (SIZE - 2 * margin - (cols - 1) * gx) // cols
    ch = 150
    top = 1430
    for i, (t, s) in enumerate(features):
        r, cc = divmod(i, cols)
        feature_card(img, margin + cc * (cw + gx), top + r * (ch + gy), cw, ch, t, s)

    pill(img, SIZE // 2, SIZE - 120,
         "15 SHEETS · INSTANT DOWNLOAD · EXCEL + GOOGLE SHEETS",
         font=fs(40), pad_x=64, pad_y=30, star=True)
    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
# 02 DASHBOARD
# ===========================================================================
def render_dashboard(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)
    page_scaffold(img, "WELLNESS DASHBOARD",
                  "Your Whole Wellbeing, at a Glance.",
                  "8 gentle KPIs + mood, energy & sleep trends — calm, never cluttered")
    box = (130, 510, SIZE - 130, SIZE - 190)
    panel(img, box, fill_color=BG)
    inner = (box[0] + 22, box[1] + 22, box[2] - 22, box[3] - 22)
    dashboard_screen(img, inner, charts_note=True)
    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
# 03 CHECK-IN
# ===========================================================================
def render_checkin(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)
    page_scaffold(img, "DAILY CHECK-IN",
                  "A Kind 30-Second Daily Ritual.",
                  "Log your mood, energy, sleep & a few notes — patterns reveal themselves")
    box = (120, 520, SIZE - 120, SIZE - 220)
    panel(img, box, title="THIS WEEK")
    inner = (box[0] + 28, box[1] + 80, box[2] - 28, box[3] - 28)
    moodcol = {7: MINT_BG, 8: MINT_BG, 6: (244, 240, 222), 5: WARN_BG}
    rows = [
        ("Mon", "8", "8", "8.0 hrs", "Rested, felt calm", MINT_BG),
        ("Tue", "7", "7", "7.5 hrs", "Good walk at lunch", None),
        ("Wed", "5", "4", "6.0 hrs", "Tough meeting — breathed through it", WARN_BG),
        ("Thu", "6", "6", "7.0 hrs", "Steady day", None),
        ("Fri", "8", "8", "8.5 hrs", "Great day outdoors", MINT_BG),
        ("Sat", "7", "7", "7.5 hrs", "Called a friend", None),
        ("Sun", "7", "7", "7.5 hrs", "Used Social Prep — went well!", MINT_BG),
    ]
    draw_table(img, inner, ["DAY", "MOOD", "ENERGY", "SLEEP", "NOTES"],
               rows, col_aligns={4: "lm"}, header_font=fs(22), row_font=fs(24))
    # small calm note
    d = ImageDraw.Draw(img)
    tc(d, (SIZE // 2, SIZE - 150), "Mood cells gently shade green when you're doing well",
       fs(26, bold=False), TEXT_MUTED)
    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
# 04 HABITS
# ===========================================================================
def render_habits(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)
    page_scaffold(img, "HABIT TRACKER",
                  "Tiny Habits, Real Calm.",
                  "Tick what you did — streaks & your weekly routine score update for you")
    box = (110, 520, SIZE - 110, SIZE - 200)
    panel(img, box, title="THIS WEEK'S HABITS")
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    inner = (box[0] + 28, box[1] + 80, box[2] - 28, box[3] - 28)
    habits = ["Meditation", "Deep Breathing", "Exercise", "Walking",
              "Healthy Meals", "Gratitude", "Limit Screens"]
    days = ["M", "T", "W", "T", "F", "S", "S"]
    done = [
        [1, 1, 0, 1, 1, 1, 1],
        [1, 1, 1, 1, 0, 1, 1],
        [0, 1, 1, 0, 1, 1, 0],
        [1, 1, 1, 1, 1, 0, 1],
        [1, 0, 1, 1, 1, 1, 1],
        [1, 1, 1, 1, 1, 1, 1],
        [0, 1, 0, 1, 1, 0, 1],
    ]
    iw = inner[2] - inner[0]
    ih = inner[3] - inner[1]
    label_w = 360
    grid_w = iw - label_w - 130   # leave room for streak col
    cell = grid_w / 7
    hdr_h = 56
    # header days
    for j, dlab in enumerate(days):
        cx = inner[0] + label_w + cell * (j + 0.5)
        od.text((cx, inner[1] + hdr_h / 2), dlab, font=fs(26), fill=ACCENT, anchor="mm")
    od.text((inner[2] - 65, inner[1] + hdr_h / 2), "★", font=fs(28), fill=GOLD, anchor="mm")
    row_h = (ih - hdr_h) / len(habits)
    for ri, hb in enumerate(habits):
        y0 = inner[1] + hdr_h + ri * row_h
        if ri % 2 == 1:
            od.rectangle((inner[0], y0, inner[2], y0 + row_h), fill=(247, 243, 235))
        od.text((inner[0] + 8, y0 + row_h / 2), hb, font=fs(26), fill=PRIMARY, anchor="lm")
        streak = 0
        for j in range(6, -1, -1):
            if done[ri][j]:
                streak += 1
            else:
                break
        for j in range(7):
            cx = inner[0] + label_w + cell * (j + 0.5)
            cyc = y0 + row_h / 2
            if done[ri][j]:
                od.ellipse((cx - 22, cyc - 22, cx + 22, cyc + 22), fill=HIGHLIGHT, outline=PRIMARY, width=3)
                od.text((cx, cyc - 1), "✓", font=fs(26), fill=PRIMARY, anchor="mm")
            else:
                od.ellipse((cx - 22, cyc - 22, cx + 22, cyc + 22), fill=BG, outline=(210, 205, 195), width=3)
        od.text((inner[2] - 65, y0 + row_h / 2), str(streak), font=fserif(32), fill=ACCENT, anchor="mm")
    img.alpha_composite(ov)
    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
# 05 PROGRESS
# ===========================================================================
def render_progress(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)
    page_scaffold(img, "PROGRESS",
                  "See How Far You've Come.",
                  "A gentle Wellness Score that blends mood, routine, sleep & journaling")

    # Wellness score ring
    cx, cy = 560, 1050
    rr = 270
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    od.ellipse((cx - rr, cy - rr, cx + rr, cy + rr), fill=WHITE, outline=(230, 224, 212), width=3)
    # arc 78%
    od.pieslice((cx - rr, cy - rr, cx + rr, cy + rr), -90, -90 + 0.78 * 360, fill=HIGHLIGHT)
    od.ellipse((cx - rr + 60, cy - rr + 60, cx + rr - 60, cy + rr - 60), fill=WHITE)
    od.text((cx, cy - 20), "78%", font=fserif(120), fill=PRIMARY, anchor="mm")
    od.text((cx, cy + 80), "WELLNESS SCORE", font=fs(28), fill=ACCENT, anchor="mm")
    img.alpha_composite(ov)
    d = ImageDraw.Draw(img)
    tc(d, (cx, cy + rr + 70), "A gentle guide, not a grade", fs(26, bold=False), TEXT_MUTED)

    # snapshot metric cards (right)
    metrics = [("Avg Mood", "6.8 / 10"), ("Avg Sleep", "7.3 hrs"),
               ("Routine", "71%"), ("Best Streak", "9 days"),
               ("Journaling", "12 days"), ("Goals Done", "1")]
    bx = 1080
    bw = SIZE - 130 - bx
    cw = (bw - 30) / 2
    chh = 200
    for i, (lab, val) in enumerate(metrics):
        r, cc = divmod(i, 2)
        x = bx + cc * (cw + 30)
        y = 780 + r * (chh + 30)
        box = (x, y, x + cw, y + chh)
        shadow(img, box, 18, 20, 70)
        ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
        od = ImageDraw.Draw(ov)
        od.rounded_rectangle(box, radius=18, fill=WHITE, outline=GOLD_LT, width=2)
        od.text((x + cw / 2, y + 56), lab, font=fs(26), fill=ACCENT, anchor="mm")
        od.text((x + cw / 2, y + 128), val, font=fserif(54), fill=PRIMARY, anchor="mm")
        img.alpha_composite(ov)
    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
# 06 MOBILE
# ===========================================================================
def render_mobile(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)
    page_scaffold(img, "WORKS EVERYWHERE",
                  "Excel · Google Sheets · Mobile",
                  "Check in from anywhere — your calm space is always with you")
    px, py = SIZE // 2, 1300
    pw, ph = 640, 1230
    phone = (px - pw // 2, py - ph // 2, px + pw // 2, py + ph // 2)
    shadow(img, phone, 64, 50, 120)
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    od.rounded_rectangle(phone, radius=64, fill=(24, 24, 28))
    bez = 22
    screen = (phone[0] + bez, phone[1] + bez + 30, phone[2] - bez, phone[3] - bez - 30)
    od.rounded_rectangle(screen, radius=44, fill=BG)
    od.rounded_rectangle((px - 95, phone[1] + 16, px + 95, phone[1] + 50), radius=18, fill=(14, 14, 18))
    sx0, sy0, sx1, sy1 = screen
    od.rounded_rectangle((sx0, sy0, sx1, sy0 + 110), radius=44, fill=PRIMARY)
    od.rectangle((sx0, sy0 + 60, sx1, sy0 + 110), fill=PRIMARY)
    od.rectangle((sx0, sy0 + 106, sx1, sy0 + 110), fill=GOLD_LT)
    od.text(((sx0 + sx1) // 2, sy0 + 56), "Calm Compass", font=fserif(38), fill=GOLD_LT, anchor="mm")
    img.alpha_composite(ov)

    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    y = sy0 + 150
    cards = [("TODAY'S MOOD", "7 / 10", PRIMARY),
             ("WELLNESS SCORE", "78%", PRIMARY),
             ("HABIT STREAK", "9 days", ACCENT),
             ("AVG SLEEP", "7.3 hrs", PRIMARY)]
    for lab, val, col in cards:
        cb = (sx0 + 30, y, sx1 - 30, y + 135)
        od.rounded_rectangle(cb, radius=18, fill=WHITE, outline=GOLD_LT, width=2)
        od.text((cb[0] + 26, y + 30), lab, font=fs(22), fill=ACCENT, anchor="lt")
        od.text((cb[0] + 26, y + 92), val, font=fserif(46), fill=col, anchor="lm")
        y += 152

    od.text((sx0 + 40, y + 16), "TODAY'S CHECK-IN", font=fs(22), fill=ACCENT, anchor="lt")
    y += 52
    for lab, state in [("Mood logged", True), ("Habits ticked", True),
                       ("Gratitude written", False), ("Wind-down planned", False)]:
        col = HIGHLIGHT if state else BG
        od.ellipse((sx0 + 40, y + 6, sx0 + 78, y + 44), fill=col, outline=PRIMARY, width=3)
        if state:
            od.text((sx0 + 59, y + 24), "✓", font=fs(24), fill=PRIMARY, anchor="mm")
        od.text((sx0 + 96, y + 24), lab, font=fs(24), fill=TEXT, anchor="lm")
        y += 64
    img.alpha_composite(ov)
    img.convert("RGB").save(out, "PNG", optimize=True)


def main():
    out_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "marketing")
    os.makedirs(out_dir, exist_ok=True)
    targets = [
        ("01_hero.png", render_hero),
        ("02_dashboard.png", render_dashboard),
        ("03_checkin.png", render_checkin),
        ("04_habits.png", render_habits),
        ("05_progress.png", render_progress),
        ("06_mobile.png", render_mobile),
    ]
    for name, fn in targets:
        fn(os.path.join(out_dir, name))
        print(f"  ✓ {name}")
    print(f"Wrote {len(targets)} images to {out_dir}")


if __name__ == "__main__":
    main()

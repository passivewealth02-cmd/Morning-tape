"""Premium feature-forward hero thumbnail for Back-to-School Command Center.

Outputs ../marketing/01_hero.png (2000x2000).
Run: python3 build_thumbnail.py
"""
from __future__ import annotations
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


def grade_crest(c, cx, cy, r=58):
    """Green roundel with an 'A+' — instantly reads 'school'."""
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    d.ellipse((cx - r, cy - r, cx + r, cy + r), fill=PRIMARY, outline=GOLD_LT, width=4)
    d.ellipse((cx - r + 10, cy - r + 10, cx + r - 10, cy + r - 10), outline=GOLD_LT, width=2)
    d.text((cx, cy), "A+", font=fserif(int(r * 0.85)), fill=GOLD_LT, anchor="mm")
    c.alpha_composite(ov)


def benefit_badge(c, cx, cy, big, small, w=440, h=170, big_color=GOLD_LT):
    box = (cx - w // 2, cy - h // 2, cx + w // 2, cy + h // 2)
    shadow(c, box, 22, 26, 90)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    d.rounded_rectangle(box, radius=22, fill=PRIMARY, outline=GOLD_LT, width=3)
    d.text((cx, cy - h * 0.20), big, font=fserif(60), fill=big_color, anchor="mm")
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


def dashboard_screen(c, screen):
    sx0, sy0, sx1, sy1 = screen
    sw = sx1 - sx0
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    head = 84
    d.rounded_rectangle((sx0, sy0, sx1, sy0 + head), radius=8, fill=PRIMARY)
    d.rectangle((sx0, sy0 + 30, sx1, sy0 + head), fill=PRIMARY)
    d.text((sx0 + 28, sy0 + head // 2), "BACK-TO-SCHOOL COMMAND CENTER", font=fs(25), fill=WHITE, anchor="lm")
    d.text((sx1 - 28, sy0 + head // 2), "2026 – 2027", font=fs(22), fill=GOLD_LT, anchor="rm")
    d.rectangle((sx0, sy0 + head, sx1, sy0 + head + 4), fill=GOLD_LT)

    gx, gy, gap = sx0 + 22, sy0 + head + 22, 14
    kw = (sw - 22 * 2 - gap * 3) // 4
    kh = 120
    kpis = [
        ("TOTAL BUDGET", "$1,200", PRIMARY),
        ("SPENT", "$558", ACCENT),
        ("REMAINING", "$642", PRIMARY),
        ("DAYS TO SCHOOL", "47", DANGER),
        ("ITEMS BOUGHT", "9", PRIMARY),
        ("ITEMS LEFT", "13", ACCENT),
        ("SUPPLY %", "41%", PRIMARY),
        ("ASSIGNMENTS", "30%", PRIMARY),
    ]
    for i, (lab, val, col) in enumerate(kpis):
        r, ci = divmod(i, 4)
        x0 = gx + ci * (kw + gap)
        y0 = gy + r * (kh + gap)
        d.rounded_rectangle((x0, y0, x0 + kw, y0 + kh), radius=10, fill=WHITE, outline=(225, 218, 205), width=2)
        d.text((x0 + 14, y0 + 16), lab, font=fs(15), fill=ACCENT, anchor="lt")
        d.text((x0 + 14, y0 + 70), val, font=fserif(38), fill=col, anchor="lm")

    nav_y = gy + 2 * (kh + gap) + 8
    nav_h = 34
    nav = ["Supplies", "Budget", "Schedule", "Assignments", "Calendar", "Meals"]
    cw = (sw - 22 * 2 - 6 * (len(nav) - 1)) / len(nav)
    for i, name in enumerate(nav):
        x0 = gx + i * (cw + 6)
        d.rounded_rectangle((x0, nav_y, x0 + cw, nav_y + nav_h), radius=17, fill=PRIMARY)
        d.text((x0 + cw / 2, nav_y + nav_h / 2), name, font=fs(15), fill=WHITE, anchor="mm")

    note_y = nav_y + nav_h + 24
    d.text((sx0 + sw / 2, note_y),
           "Supplies · Budget · Schedule · Homework · Meals · Calendar — all in one file",
           font=fs(18, bold=False), fill=TEXT_MUTED, anchor="mt")
    c.alpha_composite(ov)


def laptop(c, cx, cy, w, h):
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
    dashboard_screen(c, screen)


def render(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)
    d = ImageDraw.Draw(img)

    grade_crest(img, SIZE // 2, 150, r=58)
    pill(img, SIZE // 2, 285, "THE ULTIMATE BACK-TO-SCHOOL PLANNER", font=fs(35), pad_x=54, pad_y=22)

    tc(d, (SIZE // 2, 430), "BACK-TO-SCHOOL", fserif(128), PRIMARY)
    tc(d, (SIZE // 2, 552), "COMMAND CENTER", fserif(104), PRIMARY)
    gold_divider(img, SIZE // 2, 632, width=560)
    tc(d, (SIZE // 2, 680), "Supplies, budget, schedules & homework — all in ONE file",
       fs(26, bold=False), TEXT_MUTED)

    laptop(img, SIZE // 2, 1015, w=1360, h=560)
    benefit_badge(img, 300, 738, "47", "DAYS TO SCHOOL", big_color=GOLD_LT)
    benefit_badge(img, SIZE - 300, 738, "2-in-1", "EXCEL + GOOGLE SHEETS")

    tc(d, (SIZE // 2, 1360), "EVERYTHING FOR A STRESS-FREE YEAR", fs(34), ACCENT)
    features = [
        ("Supply Tracker", "qty · store · price"),
        ("School Budget", "9 categories"),
        ("Class Schedule", "color-coded weekly"),
        ("Homework Tracker", "overdue alerts"),
        ("Meal Planner", "+ grocery list"),
        ("School Calendar", "exams · holidays"),
        ("Clothing Inventory", "sizes + status"),
        ("Emergency Info", "printable"),
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
         "12 SHEETS · INSTANT DOWNLOAD · EXCEL + GOOGLE SHEETS",
         font=fs(40), pad_x=64, pad_y=30, star=True)
    img.convert("RGB").save(out, "PNG", optimize=True)


if __name__ == "__main__":
    out_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "marketing")
    os.makedirs(out_dir, exist_ok=True)
    render(os.path.join(out_dir, "01_hero.png"))
    print("back-to-school hero written")

"""Premium marketing image set for Home Command Center™ (6 images, 2000x2000).

  01_hero.png       - feature-forward main thumbnail
  02_dashboard.png  - executive home dashboard close-up
  03_money.png      - budget + bills
  04_homecare.png   - cleaning, chores & meals
  05_analytics.png  - household health score
  06_mobile.png     - mobile preview

Premium design: gradient backgrounds, deep-green hero band, glowing house
crest, gold-foil wordmark, gold-capped KPI cards, gradient charts.

Run: python3 build_marketing.py
"""
from __future__ import annotations
import os
from PIL import Image, ImageDraw, ImageFilter, ImageFont

PRIMARY = (27, 79, 72)
PRIMARY_DK = (18, 56, 51)
PRIMARY_LT = (33, 92, 83)
ACCENT = (147, 115, 86)
GOLD = (180, 145, 90)
GOLD_LT = (201, 168, 106)
GOLD_HI = (224, 196, 140)
SURFACE = (229, 211, 186)
HIGHLIGHT = (117, 230, 193)
BG = (251, 248, 242)
BG_TOP = (253, 250, 246)
BG_BOT = (242, 235, 223)
WHITE = (255, 255, 255)
TEXT = (51, 51, 51)
TEXT_MUTED = (130, 125, 115)
DANGER = (201, 76, 76)
MINT_BG = (227, 248, 239)
WARN_BG = (251, 240, 226)
DOT = (228, 220, 206)
SIZE = 2000

SANS_B = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
SANS_R = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
SERIF_B = "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf"


def fs(s, bold=True):
    return ImageFont.truetype(SANS_B if bold else SANS_R, s)


def fserif(s):
    return ImageFont.truetype(SERIF_B, s)


# ---------- gradient + texture ----------

def vgradient(w, h, top, bottom):
    col = Image.new("RGB", (1, h))
    px = col.load()
    for y in range(h):
        t = y / max(h - 1, 1)
        px[0, y] = tuple(int(top[i] + (bottom[i] - top[i]) * t) for i in range(3))
    return col.resize((w, h)).convert("RGBA")


def grad_round(c, box, radius, top, bottom, outline=None, width=0):
    x0, y0, x1, y1 = [int(v) for v in box]
    w, h = x1 - x0, y1 - y0
    if w <= 0 or h <= 0:
        return
    g = vgradient(w, h, top, bottom)
    mask = Image.new("L", (w, h), 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, w - 1, h - 1), radius=radius, fill=255)
    layer = Image.new("RGBA", c.size, (0, 0, 0, 0))
    layer.paste(g, (x0, y0), mask)
    c.alpha_composite(layer)
    if outline and width:
        ImageDraw.Draw(c).rounded_rectangle((x0, y0, x1, y1), radius=radius,
                                            outline=outline, width=width)


def radial_glow(c, cx, cy, r, color, strength=120):
    layer = Image.new("RGBA", c.size, (0, 0, 0, 0))
    ImageDraw.Draw(layer).ellipse((cx - r, cy - r, cx + r, cy + r), fill=color + (strength,))
    c.alpha_composite(layer.filter(ImageFilter.GaussianBlur(r // 2)))


def premium_bg(c, band_h=0):
    c.alpha_composite(vgradient(c.width, c.height, BG_TOP, BG_BOT))
    dots = Image.new("RGBA", c.size, (0, 0, 0, 0))
    dd = ImageDraw.Draw(dots)
    sp, r = 50, 3
    for y in range(sp // 2, c.height, sp):
        for x in range(sp // 2, c.width, sp):
            dd.ellipse((x - r, y - r, x + r, y + r), fill=DOT + (150,))
    c.alpha_composite(dots)
    if band_h:
        hero_band(c, band_h)


def hero_band(c, band_h):
    c.alpha_composite(vgradient(c.width, band_h, PRIMARY_LT, PRIMARY_DK), (0, 0))
    radial_glow(c, c.width // 2, band_h // 2 - 30, 520, (60, 130, 118), 70)
    wm = Image.new("RGBA", c.size, (0, 0, 0, 0))
    wd = ImageDraw.Draw(wm)
    for rr in (300, 230, 160):
        wd.ellipse((c.width - 120 - rr, band_h - 60 - rr,
                    c.width - 120 + rr, band_h - 60 + rr), outline=(255, 255, 255, 22), width=3)
    c.alpha_composite(wm)
    d = ImageDraw.Draw(c)
    d.rectangle((0, band_h - 5, c.width, band_h), fill=GOLD_LT)
    d.rectangle((0, band_h - 5, c.width, band_h - 2), fill=GOLD_HI)


# ---------- primitives ----------

def shadow(c, box, radius, blur=24, alpha=70, dy=18):
    layer = Image.new("RGBA", c.size, (0, 0, 0, 0))
    ImageDraw.Draw(layer).rounded_rectangle(box, radius=radius, fill=(18, 50, 45, alpha))
    c.alpha_composite(layer.filter(ImageFilter.GaussianBlur(blur)), (0, dy))


def tsize(d, t, f):
    b = d.textbbox((0, 0), t, font=f)
    return b[2] - b[0], b[3] - b[1]


def tc(d, xy, t, f, fill, anchor="mm"):
    d.text(xy, t, font=f, fill=fill, anchor=anchor)


def wordmark(c, cx, cy, text, size, fill=GOLD_HI, shadow_col=(8, 30, 27), max_w=None):
    d = ImageDraw.Draw(c)
    if max_w:
        while size > 20 and d.textlength(text, font=fserif(size)) > max_w:
            size -= 2
    f = fserif(size)
    sh = Image.new("RGBA", c.size, (0, 0, 0, 0))
    ImageDraw.Draw(sh).text((cx + 5, cy + 7), text, font=f, fill=shadow_col + (160,), anchor="mm")
    c.alpha_composite(sh.filter(ImageFilter.GaussianBlur(7)))
    bb = d.textbbox((cx, cy), text, font=f, anchor="mm")
    pad = 30
    bx0, by0, bx1, by1 = bb[0] - pad, bb[1] - pad, bb[2] + pad, bb[3] + pad
    w, h = bx1 - bx0, by1 - by0
    grad = vgradient(w, h, GOLD_HI, GOLD)
    mask = Image.new("L", (w, h), 0)
    ImageDraw.Draw(mask).text((cx - bx0, cy - by0), text, font=f, fill=255, anchor="mm")
    layer = Image.new("RGBA", c.size, (0, 0, 0, 0))
    layer.paste(grad, (bx0, by0), mask)
    c.alpha_composite(layer)


def pill(c, cx, cy, text, font, pad_x=60, pad_y=26, star=False, bg=PRIMARY,
         fg=WHITE, grad=None, outline=GOLD_LT):
    d = ImageDraw.Draw(c)
    label = f"★  {text}" if star else text
    tw, th = tsize(d, label, font)
    w, h = tw + pad_x * 2, th + pad_y * 2
    box = (cx - w // 2, cy - h // 2, cx + w // 2, cy + h // 2)
    shadow(c, box, h // 2, 22, 70, 12)
    if grad:
        grad_round(c, box, h // 2, grad[0], grad[1])
    else:
        ImageDraw.Draw(c).rounded_rectangle(box, radius=h // 2, fill=bg)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    od.rounded_rectangle((box[0] + 5, box[1] + 5, box[2] - 5, box[3] - 5),
                         radius=(h - 10) // 2, outline=outline, width=2)
    od.text((cx, cy), label, font=font, fill=fg, anchor="mm")
    c.alpha_composite(ov)


def gold_divider(c, cx, cy, width=560, color=GOLD):
    d = ImageDraw.Draw(c)
    d.line((cx - width // 2, cy, cx - 30, cy), fill=color, width=3)
    d.line((cx + 30, cy, cx + width // 2, cy), fill=color, width=3)
    d.polygon([(cx, cy - 12), (cx + 16, cy), (cx, cy + 12), (cx - 16, cy)], fill=color)
    d.ellipse((cx - 4, cy - 4, cx + 4, cy + 4), fill=GOLD_HI)


def house_crest(c, cx, cy, r=60, glow=True):
    if glow:
        radial_glow(c, cx, cy, int(r * 2.1), GOLD_HI, 90)
    # gradient rounded-square badge
    grad_round(c, (cx - r, cy - r, cx + r, cy + r), 22, PRIMARY_LT, PRIMARY_DK,
               outline=GOLD_LT, width=4)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    d.rounded_rectangle((cx - r + 10, cy - r + 10, cx + r - 10, cy + r - 10),
                        radius=16, outline=GOLD_HI, width=2)
    # house: roof + walls + door
    rw = r * 0.56
    roof_y = cy - r * 0.34
    eave_y = cy - r * 0.04
    d.polygon([(cx, roof_y), (cx - rw, eave_y), (cx + rw, eave_y)], fill=GOLD_HI)
    d.rectangle((cx - rw * 0.74, eave_y, cx + rw * 0.74, cy + r * 0.46), fill=HIGHLIGHT)
    d.rectangle((cx - rw * 0.16, cy + r * 0.06, cx + rw * 0.16, cy + r * 0.46), fill=PRIMARY_DK)
    d.ellipse((cx + rw * 0.04, cy + r * 0.22, cx + rw * 0.12, cy + r * 0.30), fill=GOLD_HI)
    c.alpha_composite(ov)


def stat_chip(c, cx, cy, big, small, w=400, h=150):
    box = (cx - w // 2, cy - h // 2, cx + w // 2, cy + h // 2)
    shadow(c, box, 20, 24, 80, 16)
    grad_round(c, box, 20, PRIMARY_LT, PRIMARY_DK, outline=GOLD_LT, width=3)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    d.rounded_rectangle((box[0] + 18, box[1] + 12, box[2] - 18, box[1] + 18), radius=3, fill=GOLD_HI)
    d.text((cx, cy - h * 0.16), big, font=fserif(50), fill=GOLD_HI, anchor="mm")
    d.text((cx, cy + h * 0.28), small, font=fs(22), fill=WHITE, anchor="mm")
    c.alpha_composite(ov)


def feature_card(c, x, y, w, h, title, sub):
    box = (x, y, x + w, y + h)
    shadow(c, box, 16, 18, 55, 14)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    d.rounded_rectangle(box, radius=16, fill=WHITE, outline=(232, 224, 208), width=2)
    d.rounded_rectangle((x, y, x + 8, y + h), radius=4, fill=GOLD_LT)
    d.rectangle((x + 4, y, x + 8, y + h), fill=GOLD_LT)
    c.alpha_composite(ov)
    cyc = y + h // 2
    bx = x + 58
    radial_glow(c, bx, cyc, 44, HIGHLIGHT, 70)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    d.ellipse((bx - 30, cyc - 30, bx + 30, cyc + 30), fill=HIGHLIGHT, outline=PRIMARY, width=3)
    d.text((bx, cyc - 1), "✓", font=fs(34), fill=PRIMARY, anchor="mm")
    d.text((x + 106, cyc - 22), title, font=fs(28), fill=PRIMARY, anchor="lm")
    d.text((x + 106, cyc + 24), sub, font=fs(21, bold=False), fill=TEXT_MUTED, anchor="lm")
    c.alpha_composite(ov)


def page_scaffold(img, badge, title, subtitle):
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 132, badge, font=fs(40), pad_x=56, pad_y=22, grad=(GOLD_LT, GOLD))
    tc(d, (SIZE // 2, 248), title, fserif(62), WHITE)
    gold_divider(img, SIZE // 2, 318, width=520, color=GOLD_HI)
    tc(d, (SIZE // 2, 364), subtitle, fs(30, bold=False), (226, 214, 190))


def panel(img, box, fill_color=WHITE, title=None):
    shadow(img, box, 24, 38, 95, 20)
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    od.rounded_rectangle(box, radius=24, fill=fill_color, outline=GOLD_LT, width=3)
    img.alpha_composite(ov)
    if title:
        d = ImageDraw.Draw(img)
        d.rounded_rectangle((box[0] + 34, box[1] + 26, box[0] + 42, box[1] + 56), radius=3, fill=GOLD_LT)
        d.text((box[0] + 60, box[1] + 26), title, font=fs(30), fill=ACCENT, anchor="lt")


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
    grad_round(img, (inner[0], inner[1], inner[2], inner[1] + hdr_h), 12, PRIMARY_LT, PRIMARY_DK)
    for i, h in enumerate(headers):
        cx = inner[0] + col_w * (i + 0.5)
        od.text((cx, inner[1] + hdr_h // 2), h, font=hf, fill=WHITE, anchor="mm")
    row_h = (ih - hdr_h - 10) / len(rows)
    for ri, row in enumerate(rows):
        *cells, bg = row
        y0 = inner[1] + hdr_h + ri * row_h
        if bg is None and ri % 2 == 1:
            bg = (245, 239, 228)
        if bg:
            od.rectangle((inner[0] + 3, y0, inner[2] - 3, y0 + row_h), fill=bg)
        for ci, val in enumerate(cells):
            cx = inner[0] + col_w * (ci + 0.5)
            align = (col_aligns or {}).get(ci, "mm")
            color = PRIMARY if ci == 0 else TEXT
            if status_map is not None and str(val) in status_map:
                bgc, fgc = status_map[str(val)]
                pw = min(col_w - 16, 190)
                od.rounded_rectangle((cx - pw / 2, y0 + row_h / 2 - 17, cx + pw / 2, y0 + row_h / 2 + 17),
                                     radius=16, fill=bgc)
                od.text((cx, y0 + row_h / 2), str(val), font=fs(18), fill=fgc, anchor="mm")
                continue
            if align == "lm":
                od.text((inner[0] + col_w * ci + 18, y0 + row_h / 2), str(val), font=rf, fill=color, anchor="lm")
            else:
                od.text((cx, y0 + row_h / 2), str(val), font=rf, fill=color, anchor="mm")
    img.alpha_composite(layer)


def kpi_tile(d, x0, y0, kw, kh, lab, val, col):
    d.rounded_rectangle((x0, y0, x0 + kw, y0 + kh), radius=10, fill=WHITE,
                        outline=(228, 221, 208), width=2)
    d.rounded_rectangle((x0 + 12, y0, x0 + kw - 12, y0 + 5), radius=2, fill=GOLD_LT)
    d.text((x0 + 16, y0 + 20), lab, font=fs(13), fill=ACCENT, anchor="lt")
    d.text((x0 + 16, y0 + 74), val, font=fserif(32), fill=col, anchor="lm")


def dashboard_screen(c, screen, charts_note=True):
    sx0, sy0, sx1, sy1 = screen
    sw = sx1 - sx0
    head = 84
    grad_round(c, (sx0, sy0, sx1, sy0 + head), 8, PRIMARY_LT, PRIMARY_DK)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    d.text((sx0 + 28, sy0 + head // 2), "HOME COMMAND CENTER", font=fs(25), fill=WHITE, anchor="lm")
    d.text((sx1 - 28, sy0 + head // 2), "Executive dashboard", font=fs(21), fill=GOLD_HI, anchor="rm")
    d.rectangle((sx0, sy0 + head, sx1, sy0 + head + 4), fill=GOLD_LT)
    gx, gy, gap = sx0 + 22, sy0 + head + 22, 13
    kw = (sw - 22 * 2 - gap * 3) // 4
    kh = 116
    kpis = [
        ("BUDGET LEFT", "$640", PRIMARY), ("BILLS PAID", "8 / 12", PRIMARY),
        ("CLEANING", "73%", PRIMARY), ("EVENTS / WK", "6", ACCENT),
        ("SAVINGS GOAL", "61%", PRIMARY), ("MEAL PLAN", "83%", PRIMARY),
        ("CHORE SCORE", "78%", PRIMARY), ("HOME HEALTH", "79%", PRIMARY),
    ]
    for i, (lab, val, col) in enumerate(kpis):
        r, ci = divmod(i, 4)
        x0 = gx + ci * (kw + gap)
        y0 = gy + r * (kh + gap)
        kpi_tile(d, x0, y0, kw, kh, lab, val, col)
    nav_y = gy + 2 * (kh + gap) + 8
    nav_h = 34
    nav = ["Budget", "Bills", "Meals", "Pantry", "Cleaning", "Calendar"]
    cw = (sw - 22 * 2 - 6 * (len(nav) - 1)) / len(nav)
    for i, name in enumerate(nav):
        x0 = gx + i * (cw + 6)
        d.rounded_rectangle((x0, nav_y, x0 + cw, nav_y + nav_h), radius=17, fill=PRIMARY)
        d.text((x0 + cw / 2, nav_y + nav_h / 2), name, font=fs(15), fill=WHITE, anchor="mm")
    c.alpha_composite(ov)
    cy0 = nav_y + nav_h + 22
    if charts_note and (sy1 - cy0) > 120:
        ov = Image.new("RGBA", c.size, (0, 0, 0, 0))
        d = ImageDraw.Draw(ov)
        cw2 = (sw - 22 * 2 - 14) / 2
        donuts = [
            ("MONTHLY SPENDING", [(29, PRIMARY), (14, ACCENT), (12, HIGHLIGHT), (10, SURFACE), (35, TEXT_MUTED)]),
            ("PANTRY BY CATEGORY", [(22, PRIMARY), (18, ACCENT), (16, HIGHLIGHT), (14, SURFACE), (30, (180, 90, 90))]),
        ]
        for k, (label, segs) in enumerate(donuts):
            bx = gx + k * (cw2 + 14)
            box = (bx, cy0, bx + cw2, sy1 - 18)
            d.rounded_rectangle(box, radius=10, fill=WHITE, outline=(228, 221, 208), width=2)
            d.text((box[0] + 16, box[1] + 14), label, font=fs(18), fill=ACCENT, anchor="lt")
            dcx = (box[0] + box[2]) // 2
            dcy = (box[1] + box[3]) // 2 + 18
            rr = min((box[3] - box[1]) * 0.30, cw2 * 0.20)
            ang = -90
            for pct, col in segs:
                s = pct * 3.6
                d.pieslice((dcx - rr, dcy - rr, dcx + rr, dcy + rr), ang, ang + s, fill=col)
                ang += s
            hole = rr * 0.5
            d.ellipse((dcx - hole, dcy - hole, dcx + hole, dcy + hole), fill=WHITE)
        c.alpha_composite(ov)
    else:
        ImageDraw.Draw(c).text((sx0 + sw / 2, cy0 + 6),
               "Budget · Bills · Meals · Pantry · Cleaning · Calendar — all in one file",
               font=fs(18, bold=False), fill=TEXT_MUTED, anchor="mt")


def laptop(c, cx, cy, w, h, charts_note=True):
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    sbox = (cx - w // 2, cy - h // 2 + 30, cx + w // 2, cy + h // 2 + 80)
    sh = Image.new("RGBA", c.size, (0, 0, 0, 0))
    ImageDraw.Draw(sh).rounded_rectangle(sbox, radius=40, fill=(18, 50, 45, 120))
    c.alpha_composite(sh.filter(ImageFilter.GaussianBlur(55)))
    body = (cx - w // 2, cy - h // 2, cx + w // 2, cy + h // 2)
    d.rounded_rectangle(body, radius=26, fill=(40, 42, 46), outline=(22, 22, 26), width=4)
    d.rounded_rectangle((body[0] + 6, body[1] + 6, body[2] - 6, body[1] + 60), radius=20, fill=(58, 60, 64))
    bez = 22
    screen = (body[0] + bez, body[1] + bez, body[2] - bez, body[3] - bez)
    d.rounded_rectangle(screen, radius=8, fill=BG)
    d.rounded_rectangle((cx - w // 2 - 120, cy + h // 2, cx + w // 2 + 120, cy + h // 2 + 24), radius=12, fill=(60, 62, 66))
    d.rounded_rectangle((cx - 80, cy + h // 2 + 4, cx + 80, cy + h // 2 + 14), radius=6, fill=(34, 34, 40))
    c.alpha_composite(ov)
    dashboard_screen(c, screen, charts_note=charts_note)


def render_hero(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=640)
    d = ImageDraw.Draw(img)
    house_crest(img, SIZE // 2, 150, r=60)
    pill(img, SIZE // 2, 285, "THE ULTIMATE HOUSEHOLD MANAGEMENT SYSTEM",
         font=fs(31), pad_x=50, pad_y=21, grad=(GOLD_LT, GOLD), fg=WHITE)
    wordmark(img, SIZE // 2, 445, "HOME COMMAND CENTER", 132, max_w=1780)
    gold_divider(img, SIZE // 2, 545, width=560, color=GOLD_HI)
    tc(d, (SIZE // 2, 600), "One elegant dashboard for your whole home — finances, schedules, meals & more",
       fs(27, bold=False), (224, 213, 190))
    chips = [("28", "ORGANIZED TABS"), ("2-in-1", "EXCEL + SHEETS"), ("100%", "AUTOMATED")]
    cyr = 745
    cw = 400
    total = len(chips) * cw + (len(chips) - 1) * 40
    startx = (SIZE - total) // 2 + cw // 2
    for i, (b, s) in enumerate(chips):
        stat_chip(img, startx + i * (cw + 40), cyr, b, s)
    laptop(img, SIZE // 2, 1090, w=1380, h=560, charts_note=False)
    tc(d, (SIZE // 2, 1430), "ONE COMMAND CENTER FOR EVERYTHING", fs(34), ACCENT)
    gold_divider(img, SIZE // 2, 1472, width=420, color=GOLD)
    features = [
        ("Smart Budget", "plan vs actual"),
        ("Bill Tracker", "never miss a due date"),
        ("Meal Planner", "a week in minutes"),
        ("Pantry Inventory", "auto reorder alerts"),
        ("Cleaning System", "whole-home rota"),
        ("Chore Manager", "by family member"),
        ("Home Maintenance", "seasonal reminders"),
        ("Family Calendar", "everyone in sync"),
    ]
    cols = 4
    margin = 110
    gx, gy = 24, 24
    cwf = (SIZE - 2 * margin - (cols - 1) * gx) // cols
    ch = 150
    top = 1520
    for i, (t, s) in enumerate(features):
        r, cc = divmod(i, cols)
        feature_card(img, margin + cc * (cwf + gx), top + r * (ch + gy), cwf, ch, t, s)
    pill(img, SIZE // 2, SIZE - 96,
         "28 SHEETS · INSTANT DOWNLOAD · EXCEL + GOOGLE SHEETS",
         font=fs(37), pad_x=58, pad_y=28, star=True, grad=(PRIMARY_LT, PRIMARY_DK))
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_dashboard(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=400)
    page_scaffold(img, "EXECUTIVE HOME DASHBOARD",
                  "Your Whole Home, At A Glance",
                  "10 live KPIs + spending, pantry, cleaning & goal charts — calm clarity in one screen")
    box = (130, 510, SIZE - 130, SIZE - 190)
    panel(img, box, fill_color=BG)
    inner = (box[0] + 22, box[1] + 22, box[2] - 22, box[3] - 22)
    dashboard_screen(img, inner, charts_note=True)
    pill(img, SIZE // 2, SIZE - 110, "EVERYTHING LINKED · UPDATES AUTOMATICALLY",
         font=fs(29), pad_x=50, pad_y=24, grad=(PRIMARY_LT, PRIMARY_DK))
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_money(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=400)
    page_scaffold(img, "BUDGET & BILLS",
                  "Take Control Of The Money",
                  "Plan vs actual, cash flow, savings rate — and never miss a due date again")
    kpis = [("MONTHLY INCOME", "$7,200"), ("TOTAL SPENT", "$6,560"),
            ("REMAINING", "$640"), ("SAVINGS RATE", "8%")]
    kw, kh, gap = 380, 175, 30
    sx = (SIZE - (4 * kw + 3 * gap)) // 2
    ky = 510
    for i, (lab, val) in enumerate(kpis):
        x = sx + i * (kw + gap)
        b = (x, ky, x + kw, ky + kh)
        shadow(img, b, 20, 24, 75, 16)
        ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
        od = ImageDraw.Draw(ov)
        od.rounded_rectangle(b, radius=20, fill=WHITE, outline=GOLD_LT, width=2)
        od.rounded_rectangle((x + 20, ky, x + kw - 20, ky + 6), radius=3, fill=GOLD_LT)
        od.text((x + kw // 2, ky + 56), lab, font=fs(25), fill=ACCENT, anchor="mm")
        od.text((x + kw // 2, ky + 122), val, font=fserif(54), fill=PRIMARY, anchor="mm")
        img.alpha_composite(ov)
    # spending donut
    db = (130, 760, 1000, SIZE - 180)
    panel(img, db, title="MONTHLY SPENDING")
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    dcx = (db[0] + db[2]) // 2 - 110
    dcy = (db[1] + db[3]) // 2 + 30
    rr = 210
    segs = [(28, PRIMARY, "Mortgage 28%"), (14, ACCENT, "Groceries 14%"),
            (9, HIGHLIGHT, "Childcare 9%"), (11, SURFACE, "Insurance 11%"),
            (38, (170, 150, 120), "Everything else 38%")]
    ang = -90
    for pct, col, _ in segs:
        s = pct * 3.6
        od.pieslice((dcx - rr, dcy - rr, dcx + rr, dcy + rr), ang, ang + s, fill=col)
        ang += s
    od.ellipse((dcx - 86, dcy - 86, dcx + 86, dcy + 86), fill=WHITE)
    od.text((dcx, dcy - 12), "$6.6K", font=fserif(34), fill=PRIMARY, anchor="mm")
    od.text((dcx, dcy + 28), "SPENT", font=fs(18), fill=TEXT_MUTED, anchor="mm")
    lx = dcx + rr + 38
    for i, (pct, col, lab) in enumerate(segs):
        yy = dcy - rr + 30 + i * 56
        od.rounded_rectangle((lx, yy, lx + 26, yy + 26), radius=5, fill=col)
        od.text((lx + 42, yy + 13), lab, font=fs(22), fill=TEXT, anchor="lm")
    img.alpha_composite(ov)
    # bills panel
    bb = (1060, 760, SIZE - 130, SIZE - 180)
    panel(img, bb, title="UPCOMING BILLS")
    inner = (bb[0] + 26, bb[1] + 90, bb[2] - 26, bb[3] - 30)
    smap = {"Paid": (MINT_BG, PRIMARY), "Due": (WARN_BG, ACCENT), "Overdue": ((251, 230, 230), DANGER)}
    rows = [
        ("Mortgage", "$1,850", "Paid", None),
        ("Childcare", "$600", "Due", None),
        ("Electric", "$165", "Due", None),
        ("Internet", "$75", "Paid", None),
        ("Cell Phones", "$140", "Due", None),
        ("Trash Service", "$38", "Overdue", None),
        ("Credit Card", "$420", "Due", None),
    ]
    draw_table(img, inner, ["BILL", "AMOUNT", "STATUS"], rows,
               status_map=smap, col_aligns={0: "lm"}, header_font=fs(22), row_font=fs(24))
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_homecare(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=400)
    page_scaffold(img, "CLEANING · CHORES · MEALS",
                  "Run The Home On A System",
                  "Daily to seasonal cleaning, chores by family member, and the week's meals — sorted")
    box = (120, 520, SIZE - 120, 1290)
    panel(img, box, title="CLEANING & CHORE COMMAND")
    inner = (box[0] + 28, box[1] + 80, box[2] - 28, box[3] - 28)
    smap = {"Done": (MINT_BG, PRIMARY), "In Progress": (WARN_BG, ACCENT),
            "Not Started": ((240, 236, 228), TEXT_MUTED)}
    rows = [
        ("Kitchen reset", "Mom", "Daily", "Done", None),
        ("Vacuum main floor", "Dad", "Weekly", "Done", None),
        ("Set / clear table", "Emma", "Daily", "In Progress", None),
        ("Feed the dog", "Liam", "Daily", "Done", None),
        ("Bathrooms deep wipe", "Mom", "Weekly", "Not Started", None),
        ("Fold laundry", "Emma", "Weekly", "In Progress", None),
        ("Clean fridge", "Mom", "Monthly", "Not Started", None),
    ]
    draw_table(img, inner, ["TASK", "WHO", "FREQUENCY", "STATUS"], rows,
               status_map=smap, col_aligns={0: "lm"}, header_font=fs(22), row_font=fs(24))
    # weekly meal strip
    mb = (120, 1330, SIZE - 120, SIZE - 150)
    panel(img, mb, title="THIS WEEK'S DINNERS")
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    meals = [("MON", "Sheet-pan\nchicken"), ("TUE", "Taco\nnight"), ("WED", "Spaghetti"),
             ("THU", "Salmon\n& rice"), ("FRI", "Homemade\npizza"), ("SAT", "Stir-fry"),
             ("SUN", "Pot\nroast")]
    cw = (mb[2] - mb[0] - 56) / 7
    for i, (day, meal) in enumerate(meals):
        x = mb[0] + 28 + i * cw
        cardb = (x + 6, mb[1] + 96, x + cw - 6, mb[3] - 28)
        grad_round(img, (cardb[0], cardb[1], cardb[2], cardb[1] + 46), 10, PRIMARY_LT, PRIMARY_DK)
        od.rounded_rectangle((cardb[0], cardb[1] + 40, cardb[2], cardb[3]), radius=10,
                             fill=WHITE, outline=(228, 221, 208), width=2)
        od.text(((cardb[0] + cardb[2]) / 2, cardb[1] + 23), day, font=fs(20), fill=GOLD_HI, anchor="mm")
        od.multiline_text(((cardb[0] + cardb[2]) / 2, (cardb[1] + 40 + cardb[3]) / 2), meal,
                          font=fs(21), fill=PRIMARY, anchor="mm", align="center", spacing=6)
    img.alpha_composite(ov)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_analytics(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=400)
    page_scaffold(img, "HOUSEHOLD HEALTH SCORE",
                  "See How Your Home Is Running",
                  "One blended score across money, meals, cleaning, chores & goals — plus the details")
    cx, cy = 560, 1050
    rr = 270
    radial_glow(img, cx, cy, 330, HIGHLIGHT, 60)
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    od.ellipse((cx - rr, cy - rr, cx + rr, cy + rr), fill=WHITE, outline=(232, 226, 214), width=3)
    od.pieslice((cx - rr, cy - rr, cx + rr, cy + rr), -90, -90 + 0.79 * 360, fill=HIGHLIGHT)
    od.ellipse((cx - rr + 60, cy - rr + 60, cx + rr - 60, cy + rr - 60), fill=WHITE)
    od.text((cx, cy - 20), "79%", font=fserif(120), fill=PRIMARY, anchor="mm")
    od.text((cx, cy + 80), "HOME HEALTH", font=fs(26), fill=ACCENT, anchor="mm")
    img.alpha_composite(ov)
    bars = [("Bills Paid", 0.67), ("Cleaning Done", 0.73), ("Chore Completion", 0.78),
            ("Savings Funded", 0.61), ("Goal Progress", 0.45), ("Meal Plan", 0.83)]
    bx = 1080
    bw = SIZE - 130 - bx
    y = 760
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    for lab, pct in bars:
        od.text((bx, y), lab, font=fs(27), fill=PRIMARY, anchor="lt")
        od.rounded_rectangle((bx, y + 44, bx + bw, y + 86), radius=21, fill=(238, 232, 222))
        grad_round(img, (bx, y + 44, bx + bw * pct, y + 86), 21, HIGHLIGHT, (70, 200, 165))
        od.text((bx + bw - 10, y + 6), f"{int(pct*100)}%", font=fs(25), fill=ACCENT, anchor="rt")
        y += 132
    img.alpha_composite(ov)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_mobile(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=400)
    page_scaffold(img, "WORKS EVERYWHERE",
                  "Excel · Google Sheets · Mobile",
                  "Update from the kitchen, the car, or the couch — your home runs from your pocket")
    px, py = SIZE // 2, 1300
    pw, ph = 640, 1230
    phone = (px - pw // 2, py - ph // 2, px + pw // 2, py + ph // 2)
    shadow(img, phone, 64, 50, 110, 24)
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    od.rounded_rectangle(phone, radius=64, fill=(26, 26, 30))
    bez = 22
    screen = (phone[0] + bez, phone[1] + bez + 30, phone[2] - bez, phone[3] - bez - 30)
    od.rounded_rectangle(screen, radius=44, fill=BG)
    od.rounded_rectangle((px - 95, phone[1] + 16, px + 95, phone[1] + 50), radius=18, fill=(14, 14, 18))
    img.alpha_composite(ov)
    sx0, sy0, sx1, sy1 = screen
    grad_round(img, (sx0, sy0, sx1, sy0 + 110), 44, PRIMARY_LT, PRIMARY_DK)
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    od.rectangle((sx0, sy0 + 106, sx1, sy0 + 110), fill=GOLD_LT)
    od.text(((sx0 + sx1) // 2, sy0 + 56), "Home Command", font=fserif(36), fill=GOLD_HI, anchor="mm")
    y = sy0 + 150
    cards = [("BUDGET LEFT", "$640", PRIMARY), ("HOME HEALTH", "79%", PRIMARY),
             ("BILLS PAID", "8 / 12", ACCENT), ("EVENTS THIS WEEK", "6", PRIMARY)]
    for lab, val, col in cards:
        cb = (sx0 + 30, y, sx1 - 30, y + 135)
        od.rounded_rectangle(cb, radius=18, fill=WHITE, outline=GOLD_LT, width=2)
        od.rounded_rectangle((cb[0] + 20, y, cb[2] - 20, y + 5), radius=2, fill=GOLD_LT)
        od.text((cb[0] + 26, y + 34), lab, font=fs(22), fill=ACCENT, anchor="lt")
        od.text((cb[0] + 26, y + 94), val, font=fserif(46), fill=col, anchor="lm")
        y += 152
    od.text((sx0 + 40, y + 16), "TODAY'S FOCUS", font=fs(22), fill=ACCENT, anchor="lt")
    y += 52
    for lab, state in [("Pay electric bill", True), ("Meal-prep dinner", True),
                       ("Soccer pickup 4pm", False), ("Restock pantry", False)]:
        col = HIGHLIGHT if state else BG
        od.ellipse((sx0 + 40, y + 6, sx0 + 78, y + 44), fill=col, outline=PRIMARY, width=3)
        if state:
            od.text((sx0 + 59, y + 24), "✓", font=fs(24), fill=PRIMARY, anchor="mm")
        od.text((sx0 + 96, y + 24), lab, font=fs(23), fill=TEXT, anchor="lm")
        y += 64
    img.alpha_composite(ov)
    img.convert("RGB").save(out, "PNG", optimize=True)


def main():
    out_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "marketing")
    os.makedirs(out_dir, exist_ok=True)
    targets = [
        ("01_hero.png", render_hero),
        ("02_dashboard.png", render_dashboard),
        ("03_money.png", render_money),
        ("04_homecare.png", render_homecare),
        ("05_analytics.png", render_analytics),
        ("06_mobile.png", render_mobile),
    ]
    for name, fn in targets:
        fn(os.path.join(out_dir, name))
        print(f"  ✓ {name}")
    print(f"Wrote {len(targets)} images to {out_dir}")


if __name__ == "__main__":
    main()

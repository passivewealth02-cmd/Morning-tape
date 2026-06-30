"""Marketing image set for Home Command Center™ (6 images, 2000x2000).

These are dense, authentic "app screenshots" — a left sidebar listing all 28
real tabs, KPI cards with the REAL computed numbers from the workbook's
sample data, and fully populated tables/charts. Built to look exactly like
the product and to sell its depth.

  01_hero.png            - branded hero + live dashboard window
  02_inside.png          - "everything inside — 28 powerful tabs"
  03_dashboard.png       - full executive dashboard screenshot
  04_budget_bills.png    - budget table + bills screenshot
  05_meals_cleaning.png  - meal planner + cleaning & chores screenshot
  06_mobile.png          - mobile preview

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
MINT = (117, 230, 193)
BG = (251, 248, 242)
BG_TOP = (253, 250, 246)
BG_BOT = (242, 235, 223)
WHITE = (255, 255, 255)
TEXT = (51, 51, 51)
TEXT_MUTED = (132, 126, 116)
DANGER = (201, 76, 76)
MINT_BG = (227, 248, 239)
WARN_BG = (251, 240, 226)
RED_BG = (251, 230, 230)
GRID = (228, 222, 210)
ROW_ALT = (246, 241, 232)
DOT = (228, 220, 206)
SIZE = 2000

SANS_B = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
SANS_R = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
SERIF_B = "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf"

# real tab list (28, excludes the bonus Welcome page)
TABS = ["Dashboard", "Family Directory", "Calendar", "Daily Command", "Budget",
        "Bills", "Grocery", "Pantry", "Fridge & Freezer", "Meal Planner",
        "Recipes", "Cleaning", "Chores", "Maintenance", "Home Inventory",
        "Subscriptions", "Family Goals", "Savings", "Holidays & Events",
        "Travel", "Children's Hub", "Pet Care", "Wellness", "Projects",
        "Documents", "Analytics", "Inspiration", "Settings"]


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
        ImageDraw.Draw(c).rounded_rectangle((x0, y0, x1, y1), radius=radius, outline=outline, width=width)


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
            dd.ellipse((x - r, y - r, x + r, y + r), fill=DOT + (140,))
    c.alpha_composite(dots)
    if band_h:
        hero_band(c, band_h)


def hero_band(c, band_h):
    c.alpha_composite(vgradient(c.width, band_h, PRIMARY_LT, PRIMARY_DK), (0, 0))
    radial_glow(c, c.width // 2, band_h // 2 - 30, 520, (60, 130, 118), 70)
    wm = Image.new("RGBA", c.size, (0, 0, 0, 0))
    wd = ImageDraw.Draw(wm)
    for rr in (300, 230, 160):
        wd.ellipse((c.width - 120 - rr, band_h - 60 - rr, c.width - 120 + rr, band_h - 60 + rr),
                   outline=(255, 255, 255, 22), width=3)
    c.alpha_composite(wm)
    d = ImageDraw.Draw(c)
    d.rectangle((0, band_h - 5, c.width, band_h), fill=GOLD_LT)
    d.rectangle((0, band_h - 5, c.width, band_h - 2), fill=GOLD_HI)


# ---------- primitives ----------

def shadow(c, box, radius, blur=24, alpha=70, dy=18):
    layer = Image.new("RGBA", c.size, (0, 0, 0, 0))
    ImageDraw.Draw(layer).rounded_rectangle(box, radius=radius, fill=(18, 50, 45, alpha))
    c.alpha_composite(layer.filter(ImageFilter.GaussianBlur(blur)), (0, dy))


def tc(d, xy, t, f, fill, anchor="mm"):
    d.text(xy, t, font=f, fill=fill, anchor=anchor)


def wordmark(c, cx, cy, text, size, fill=GOLD_HI, max_w=None):
    d = ImageDraw.Draw(c)
    if max_w:
        while size > 20 and d.textlength(text, font=fserif(size)) > max_w:
            size -= 2
    f = fserif(size)
    sh = Image.new("RGBA", c.size, (0, 0, 0, 0))
    ImageDraw.Draw(sh).text((cx + 5, cy + 7), text, font=f, fill=(8, 30, 27, 160), anchor="mm")
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


def pill(c, cx, cy, text, font, pad_x=60, pad_y=26, star=False, fg=WHITE, grad=(GOLD_LT, GOLD), outline=GOLD_HI):
    d = ImageDraw.Draw(c)
    label = f"★  {text}" if star else text
    tw = d.textlength(label, font=font)
    th = font.size
    w, h = tw + pad_x * 2, th + pad_y * 2
    box = (cx - w // 2, cy - h // 2, cx + w // 2, cy + h // 2)
    shadow(c, box, h // 2, 22, 70, 12)
    grad_round(c, box, h // 2, grad[0], grad[1])
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    od.rounded_rectangle((box[0] + 5, box[1] + 5, box[2] - 5, box[3] - 5), radius=(h - 10) // 2, outline=outline, width=2)
    od.text((cx, cy), label, font=font, fill=fg, anchor="mm")
    c.alpha_composite(ov)


def gold_divider(c, cx, cy, width=560, color=GOLD_HI):
    d = ImageDraw.Draw(c)
    d.line((cx - width // 2, cy, cx - 30, cy), fill=color, width=3)
    d.line((cx + 30, cy, cx + width // 2, cy), fill=color, width=3)
    d.polygon([(cx, cy - 12), (cx + 16, cy), (cx, cy + 12), (cx - 16, cy)], fill=color)


def house_crest(c, cx, cy, r=60, glow=True):
    if glow:
        radial_glow(c, cx, cy, int(r * 2.1), GOLD_HI, 90)
    grad_round(c, (cx - r, cy - r, cx + r, cy + r), 22, PRIMARY_LT, PRIMARY_DK, outline=GOLD_LT, width=4)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    d.rounded_rectangle((cx - r + 10, cy - r + 10, cx + r - 10, cy + r - 10), radius=16, outline=GOLD_HI, width=2)
    rw = r * 0.56
    roof_y = cy - r * 0.34
    eave_y = cy - r * 0.04
    d.polygon([(cx, roof_y), (cx - rw, eave_y), (cx + rw, eave_y)], fill=GOLD_HI)
    d.rectangle((cx - rw * 0.74, eave_y, cx + rw * 0.74, cy + r * 0.46), fill=HIGHLIGHT)
    d.rectangle((cx - rw * 0.16, cy + r * 0.06, cx + rw * 0.16, cy + r * 0.46), fill=PRIMARY_DK)
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


# ---------- chart primitives ----------

def donut(d, cx, cy, r, segs, center_top=None, center_sub=None, hole=0.55):
    ang = -90
    for pct, col in segs:
        s = pct * 3.6
        d.pieslice((cx - r, cy - r, cx + r, cy + r), ang, ang + s, fill=col)
        ang += s
    hr = r * hole
    d.ellipse((cx - hr, cy - hr, cx + hr, cy + hr), fill=WHITE)
    if center_top:
        d.text((cx, cy - (10 if center_sub else 0)), center_top, font=fserif(int(r * 0.34)), fill=PRIMARY, anchor="mm")
    if center_sub:
        d.text((cx, cy + int(r * 0.26)), center_sub, font=fs(int(r * 0.13)), fill=TEXT_MUTED, anchor="mm")


def legend(d, x, y, items, fsz=20, gap=40):
    for i, (col, lab) in enumerate(items):
        yy = y + i * gap
        d.rounded_rectangle((x, yy, x + 24, yy + 24), radius=5, fill=col)
        d.text((x + 36, yy + 12), lab, font=fs(fsz, bold=False), fill=TEXT, anchor="lm")


def vbars(img, d, box, items, maxv=1.0, fmt_pct=True):
    x0, y0, x1, y1 = box
    n = len(items)
    bw = (x1 - x0) / n
    base = y1 - 36
    for i, (lab, val, col) in enumerate(items):
        cx = x0 + bw * (i + 0.5)
        h = (val / maxv) * (base - y0 - 10)
        grad_round(img, (cx - bw * 0.3, base - h, cx + bw * 0.3, base), 8, PRIMARY_LT if col is None else col[0], PRIMARY_DK if col is None else col[1])
        lab_txt = f"{int(val*100)}%" if fmt_pct else str(val)
        d.text((cx, base - h - 22), lab_txt, font=fs(19), fill=PRIMARY, anchor="mm")
        d.text((cx, base + 18), lab, font=fs(18, bold=False), fill=TEXT_MUTED, anchor="mm")


def hbars(img, d, box, items):
    x0, y0, x1, y1 = box
    n = len(items)
    rowh = (y1 - y0) / n
    bw_max = (x1 - x0) - 360
    for i, (lab, pct) in enumerate(items):
        yy = y0 + rowh * i + rowh * 0.2
        d.text((x0, yy), lab, font=fs(21, bold=False), fill=TEXT, anchor="lt")
        track = (x0, yy + 34, x0 + bw_max, yy + 64)
        d.rounded_rectangle(track, radius=15, fill=(236, 230, 220))
        grad_round(img, (x0, yy + 34, x0 + bw_max * pct, yy + 64), 15, HIGHLIGHT, (70, 200, 165))
        d.text((x0 + bw_max + 16, yy + 49), f"{int(pct*100)}%", font=fs(21), fill=ACCENT, anchor="lm")


# ---------- app window chrome ----------

def app_window(img, box, active_idx, content_fn, file_label="Home_Command_Center.xlsx — The Anderson Home"):
    x0, y0, x1, y1 = box
    shadow(img, box, 26, 40, 95, 22)
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    od.rounded_rectangle(box, radius=24, fill=WHITE, outline=(210, 203, 190), width=2)
    img.alpha_composite(ov)
    # title bar
    tb_h = 58
    grad_round(img, (x0, y0, x1, y0 + tb_h + 24), 24, (54, 56, 60), (44, 46, 50))
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    od.rectangle((x0, y0 + tb_h, x1, y0 + tb_h + 4), fill=(36, 38, 42))
    for i, col in enumerate([(237, 106, 94), (245, 191, 79), (98, 197, 84)]):
        od.ellipse((x0 + 30 + i * 36, y0 + tb_h // 2 - 11, x0 + 52 + i * 36, y0 + tb_h // 2 + 11), fill=col)
    od.text(((x0 + x1) / 2, y0 + tb_h // 2), file_label, font=fs(22, bold=False), fill=(225, 222, 215), anchor="mm")
    img.alpha_composite(ov)
    # sidebar
    sb_w = int((x1 - x0) * 0.205)
    sb = (x0, y0 + tb_h, x0 + sb_w, y1)
    grad_round(img, (sb[0], sb[1], sb[2] + 24, sb[3]), 0, PRIMARY_LT, PRIMARY_DK)
    grad_round(img, (sb[0], y1 - 24, sb[2], y1), 24, PRIMARY_DK, PRIMARY_DK)  # bottom-left round
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    # sidebar brand
    bx = sb[0] + 26
    od.text((bx, sb[1] + 30), "HOME COMMAND", font=fs(20), fill=GOLD_HI, anchor="lt")
    od.text((bx, sb[1] + 58), "28-tab system", font=fs(15, bold=False), fill=(170, 200, 192), anchor="lt")
    od.line((sb[0] + 20, sb[1] + 92, sb[2] - 16, sb[1] + 92), fill=(255, 255, 255, 40), width=1)
    # tab list
    list_top = sb[1] + 108
    rowh = (y1 - 24 - list_top) / len(TABS)
    palette = [HIGHLIGHT, GOLD_HI, SURFACE, (150, 200, 190)]
    for i, name in enumerate(TABS):
        ry = list_top + i * rowh
        if i == active_idx:
            od.rounded_rectangle((sb[0] + 12, ry + 2, sb[2] - 10, ry + rowh - 2), radius=8, fill=(255, 255, 255, 235))
            od.rounded_rectangle((sb[0] + 12, ry + 2, sb[0] + 19, ry + rowh - 2), radius=3, fill=GOLD_HI)
            dotc = PRIMARY
            txtc = PRIMARY
            font = fs(19)
        else:
            dotc = palette[i % len(palette)]
            txtc = (214, 226, 222)
            font = fs(18, bold=False)
        cyr = ry + rowh / 2
        od.ellipse((sb[0] + 30, cyr - 6, sb[0] + 42, cyr + 6), fill=dotc)
        od.text((sb[0] + 56, cyr), name, font=font, fill=txtc, anchor="lm")
    img.alpha_composite(ov)
    # content
    cbox = (sb[2] + 1, y0 + tb_h + 4, x1, y1)
    content_fn(img, cbox)


# ---------- content renderers (real data) ----------

KPIS = [
    ("BUDGET LEFT", "$223", "of $6,450 planned"),
    ("BILLS PAID", "8 / 12", "1 overdue"),
    ("CLEANING DONE", "5 / 15", "tasks complete"),
    ("PANTRY STOCK", "50%", "6 items low"),
    ("EVENTS / WEEK", "6", "this week"),
    ("MEAL PLAN", "91%", "32 / 35 slots"),
    ("SAVINGS GOAL", "32%", "$29.1K / $90.5K"),
    ("PROJECTS", "4", "active now"),
    ("MAINT. DUE", "5", "next 30 days"),
    ("HOME HEALTH", "62%", "overall score"),
]


def content_dashboard(img, cbox):
    x0, y0, x1, y1 = cbox
    pad = 30
    d = ImageDraw.Draw(img)
    d.rectangle(cbox, fill=BG)
    # header
    d.text((x0 + pad, y0 + 24), "Executive Home Dashboard", font=fs(34), fill=PRIMARY, anchor="lt")
    d.text((x0 + pad, y0 + 66), "The Anderson Home  ·  live KPIs update as you type", font=fs(20, bold=False), fill=TEXT_MUTED, anchor="lt")
    d.rounded_rectangle((x1 - pad - 150, y0 + 28, x1 - pad, y0 + 64), radius=18, fill=MINT_BG)
    d.text((x1 - pad - 75, y0 + 46), "● live", font=fs(18), fill=PRIMARY, anchor="mm")
    # KPI grid 5 x 2
    gx = x0 + pad
    gy = y0 + 100
    gw = (x1 - x0 - 2 * pad)
    gap = 16
    kw = (gw - 4 * gap) / 5
    kh = 124
    for i, (lab, val, sub) in enumerate(KPIS):
        r, ci = divmod(i, 5)
        kx = gx + ci * (kw + gap)
        ky = gy + r * (kh + gap)
        d.rounded_rectangle((kx, ky, kx + kw, ky + kh), radius=12, fill=WHITE, outline=GRID, width=2)
        d.rounded_rectangle((kx + 14, ky, kx + kw - 14, ky + 5), radius=2, fill=GOLD_LT)
        d.text((kx + 16, ky + 18), lab, font=fs(14), fill=ACCENT, anchor="lt")
        d.text((kx + 16, ky + 62), val, font=fserif(34), fill=PRIMARY, anchor="lm")
        d.text((kx + 16, ky + 100), sub, font=fs(14, bold=False), fill=TEXT_MUTED, anchor="lm")
    # charts band
    cy_top = gy + 2 * (kh + gap) + 20
    d.text((gx, cy_top), "MONEY & HOME AT A GLANCE", font=fs(20), fill=ACCENT, anchor="lt")
    panels_y = cy_top + 36
    panel_h = (y1 - panels_y - pad)
    pw = (gw - 3 * gap) / 4
    # 1 spending donut
    cards = [
        ("Monthly Spending", "donut", [(28, PRIMARY), (14, ACCENT), (9, HIGHLIGHT), (11, SURFACE), (38, (170, 150, 120))],
         ("$6.2K", "spent"), [(PRIMARY, "Mortgage 28%"), (ACCENT, "Groceries 14%"), ((170,150,120), "Other 38%")]),
        ("Bill Status", "donut", [(67, PRIMARY), (33, (210, 150, 90))], ("8/12", "paid"),
         [(PRIMARY, "Paid 8"), ((210, 150, 90), "Unpaid 4")]),
    ]
    for k, (title, _, segs, center, leg) in enumerate(cards):
        px = gx + k * (pw + gap)
        pb = (px, panels_y, px + pw, panels_y + panel_h)
        d.rounded_rectangle(pb, radius=12, fill=WHITE, outline=GRID, width=2)
        d.text((px + 16, panels_y + 14), title, font=fs(17), fill=ACCENT, anchor="lt")
        dcx, dcy = px + pw * 0.42, panels_y + panel_h * 0.55
        rr = min(panel_h * 0.30, pw * 0.28)
        donut(d, dcx, dcy, rr, segs, center[0], center[1])
        legend(d, px + pw * 0.04, panels_y + panel_h - 110, leg, fsz=15, gap=32)
    # 3 cleaning bars
    px = gx + 2 * (pw + gap)
    pb = (px, panels_y, px + pw, panels_y + panel_h)
    d.rounded_rectangle(pb, radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Cleaning Progress", font=fs(17), fill=ACCENT, anchor="lt")
    vbars(img, d, (px + 24, panels_y + 50, px + pw - 16, panels_y + panel_h - 6),
          [("Daily", 1.0, None), ("Wk", 0.4, None), ("Mo", 0.0, None), ("Seas", 0.0, None), ("Deep", 0.0, None)])
    # 4 chore bars by member
    px = gx + 3 * (pw + gap)
    pb = (px, panels_y, px + pw, panels_y + panel_h)
    d.rounded_rectangle(pb, radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Chore Completion", font=fs(17), fill=ACCENT, anchor="lt")
    vbars(img, d, (px + 24, panels_y + 50, px + pw - 16, panels_y + panel_h - 6),
          [("Mom", 1.0, (HIGHLIGHT, (70, 200, 165))), ("Dad", 1.0, (HIGHLIGHT, (70, 200, 165))),
           ("Emma", 0.78, (HIGHLIGHT, (70, 200, 165))), ("Liam", 0.57, (HIGHLIGHT, (70, 200, 165)))])


def content_budget(img, cbox):
    x0, y0, x1, y1 = cbox
    pad = 30
    d = ImageDraw.Draw(img)
    d.rectangle(cbox, fill=BG)
    d.text((x0 + pad, y0 + 22), "Household Budget", font=fs(32), fill=PRIMARY, anchor="lt")
    d.text((x0 + pad, y0 + 62), "Plan vs actual  ·  cash flow & savings rate, automatic", font=fs(19, bold=False), fill=TEXT_MUTED, anchor="lt")
    # KPI strip
    strip = [("INCOME", "$7,200"), ("SPENT", "$6,227"), ("REMAINING", "$223"),
             ("CASH FLOW", "+$973"), ("SAVINGS RATE", "8%")]
    sy = y0 + 96
    sw = (x1 - x0 - 2 * pad - 4 * 14) / 5
    for i, (lab, val) in enumerate(strip):
        sx = x0 + pad + i * (sw + 14)
        d.rounded_rectangle((sx, sy, sx + sw, sy + 96), radius=10, fill=WHITE, outline=GRID, width=2)
        d.rounded_rectangle((sx + 12, sy, sx + sw - 12, sy + 5), radius=2, fill=GOLD_LT)
        d.text((sx + sw / 2, sy + 30), lab, font=fs(15), fill=ACCENT, anchor="mm")
        d.text((sx + sw / 2, sy + 66), val, font=fserif(30), fill=PRIMARY, anchor="mm")
    # table
    rows = [
        ("Mortgage/Rent", 1850, 1850), ("Utilities", 300, 318), ("Internet", 75, 75),
        ("Phones", 140, 140), ("Insurance", 360, 360), ("Groceries", 850, 902),
        ("Fuel", 220, 205), ("Transportation", 180, 165), ("Childcare", 600, 600),
        ("Clothing", 150, 88), ("Entertainment", 200, 175), ("Medical", 160, 60),
        ("Savings", 600, 600), ("Subscriptions", 90, 96), ("Dining Out", 240, 268),
        ("Home Improvement", 200, 120), ("Pets", 95, 110), ("Miscellaneous", 140, 95),
    ]
    ty = sy + 124
    tx0 = x0 + pad
    tx1 = x1 - pad
    headers = ["CATEGORY", "PLANNED", "ACTUAL", "REMAINING", "% USED"]
    colx = [tx0, tx0 + (tx1 - tx0) * 0.40, tx0 + (tx1 - tx0) * 0.55, tx0 + (tx1 - tx0) * 0.70, tx0 + (tx1 - tx0) * 0.85]
    hdr_h = 46
    grad_round(img, (tx0, ty, tx1, ty + hdr_h), 8, PRIMARY_LT, PRIMARY_DK)
    for i, h in enumerate(headers):
        anc = "lm" if i == 0 else "mm"
        hx = colx[i] + (14 if i == 0 else (tx1 - tx0) * 0.075)
        d.text((hx, ty + hdr_h / 2), h, font=fs(17), fill=WHITE, anchor=anc)
    rh = (y1 - pad - (ty + hdr_h)) / (len(rows) + 1)
    barw = (tx1 - tx0) * 0.13
    for i, (cat, pl, ac) in enumerate(rows):
        ry = ty + hdr_h + i * rh
        if i % 2:
            d.rectangle((tx0, ry, tx1, ry + rh), fill=ROW_ALT)
        d.text((tx0 + 14, ry + rh / 2), cat, font=fs(18, bold=False), fill=TEXT, anchor="lm")
        d.text((colx[1] + (tx1 - tx0) * 0.075, ry + rh / 2), f"${pl:,}", font=fs(18, bold=False), fill=TEXT, anchor="mm")
        d.text((colx[2] + (tx1 - tx0) * 0.075, ry + rh / 2), f"${ac:,}", font=fs(18), fill=PRIMARY, anchor="mm")
        rem = pl - ac
        rc = DANGER if rem < 0 else TEXT
        d.text((colx[3] + (tx1 - tx0) * 0.075, ry + rh / 2), f"${rem:,}", font=fs(18, bold=False), fill=rc, anchor="mm")
        used = ac / pl if pl else 0
        tbx0 = colx[4] + 6
        d.rounded_rectangle((tbx0, ry + rh / 2 - 11, tbx0 + barw, ry + rh / 2 + 11), radius=10, fill=(236, 230, 220))
        fillw = barw * min(used, 1.0)
        grad_round(img, (tbx0, ry + rh / 2 - 11, tbx0 + fillw, ry + rh / 2 + 11), 10,
                   PRIMARY_LT if used <= 1 else DANGER, PRIMARY_DK if used <= 1 else (160, 50, 50))
        d.text((tbx0 + barw + 14, ry + rh / 2), f"{int(used*100)}%", font=fs(16), fill=ACCENT, anchor="lm")
    # totals row
    ry = ty + hdr_h + len(rows) * rh
    d.rectangle((tx0, ry, tx1, ry + rh), fill=SURFACE)
    d.text((tx0 + 14, ry + rh / 2), "TOTAL", font=fs(19), fill=PRIMARY, anchor="lm")
    d.text((colx[1] + (tx1 - tx0) * 0.075, ry + rh / 2), "$6,450", font=fs(19), fill=PRIMARY, anchor="mm")
    d.text((colx[2] + (tx1 - tx0) * 0.075, ry + rh / 2), "$6,227", font=fs(19), fill=PRIMARY, anchor="mm")
    d.text((colx[3] + (tx1 - tx0) * 0.075, ry + rh / 2), "$223", font=fs(19), fill=PRIMARY, anchor="mm")
    d.text((colx[4] + 6 + barw + 14, ry + rh / 2), "97%", font=fs(17), fill=PRIMARY, anchor="lm")


def content_bills(img, cbox):
    x0, y0, x1, y1 = cbox
    pad = 30
    d = ImageDraw.Draw(img)
    d.rectangle(cbox, fill=BG)
    d.text((x0 + pad, y0 + 22), "Bill Command Center", font=fs(32), fill=PRIMARY, anchor="lt")
    d.text((x0 + pad, y0 + 62), "Overdue bills flag themselves — never miss a due date", font=fs(19, bold=False), fill=TEXT_MUTED, anchor="lt")
    rows = [
        ("Mortgage", "Jul 2", "$1,850", "Paid"), ("Electric", "Jul 6", "$165", "Paid"),
        ("Water/Sewer", "Jul 11", "$78", "Due"), ("Internet", "Jul 8", "$75", "Paid"),
        ("Cell Phones", "Jul 14", "$140", "Paid"), ("Home Insurance", "Jul 20", "$180", "Paid"),
        ("Auto Insurance", "Jul 20", "$180", "Paid"), ("Childcare", "Jul 1", "$600", "Due"),
        ("Trash Service", "Jun 28", "$38", "Overdue"), ("Gym Membership", "Jul 16", "$45", "Paid"),
        ("Streaming Bundle", "Jul 9", "$38", "Paid"), ("Credit Card", "Jul 4", "$420", "Due"),
    ]
    headers = ["BILL", "DUE", "AMOUNT", "STATUS"]
    tx0, tx1 = x0 + pad, x1 - pad
    ty = y0 + 104
    colx = [tx0, tx0 + (tx1 - tx0) * 0.46, tx0 + (tx1 - tx0) * 0.64, tx0 + (tx1 - tx0) * 0.82]
    hdr_h = 46
    grad_round(img, (tx0, ty, tx1, ty + hdr_h), 8, PRIMARY_LT, PRIMARY_DK)
    for i, h in enumerate(headers):
        anc = "lm" if i == 0 else "mm"
        hx = colx[i] + (14 if i == 0 else (tx1 - tx0) * 0.09)
        d.text((hx, ty + hdr_h / 2), h, font=fs(17), fill=WHITE, anchor=anc)
    rh = (y1 - pad - (ty + hdr_h)) / len(rows)
    smap = {"Paid": (MINT_BG, PRIMARY), "Due": (WARN_BG, ACCENT), "Overdue": (RED_BG, DANGER)}
    for i, (name, due, amt, st) in enumerate(rows):
        ry = ty + hdr_h + i * rh
        if st == "Overdue":
            d.rectangle((tx0, ry, tx1, ry + rh), fill=(252, 240, 240))
        elif i % 2:
            d.rectangle((tx0, ry, tx1, ry + rh), fill=ROW_ALT)
        d.text((tx0 + 14, ry + rh / 2), name, font=fs(19, bold=False), fill=TEXT, anchor="lm")
        d.text((colx[1] + (tx1 - tx0) * 0.09, ry + rh / 2), due, font=fs(18, bold=False), fill=TEXT_MUTED, anchor="mm")
        d.text((colx[2] + (tx1 - tx0) * 0.09, ry + rh / 2), amt, font=fs(19), fill=PRIMARY, anchor="mm")
        bg, fg = smap[st]
        cxp = colx[3] + (tx1 - tx0) * 0.09
        d.rounded_rectangle((cxp - 70, ry + rh / 2 - 17, cxp + 70, ry + rh / 2 + 17), radius=16, fill=bg)
        d.text((cxp, ry + rh / 2), st, font=fs(17), fill=fg, anchor="mm")


def content_meals(img, cbox):
    x0, y0, x1, y1 = cbox
    pad = 30
    d = ImageDraw.Draw(img)
    d.rectangle(cbox, fill=BG)
    d.text((x0 + pad, y0 + 22), "Meal Planner", font=fs(32), fill=PRIMARY, anchor="lt")
    d.text((x0 + pad, y0 + 62), "A full week's menu in minutes  ·  91% planned", font=fs(19, bold=False), fill=TEXT_MUTED, anchor="lt")
    days = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"]
    meals = ["Breakfast", "Lunch", "Dinner", "Snack", "Dessert"]
    menu = {
        "Breakfast": ["Oatmeal", "Eggs & toast", "Smoothies", "Pancakes", "Yogurt", "Waffles", "Bagels"],
        "Lunch": ["Turkey wraps", "Lftvr chili", "Pasta salad", "Grilled cheese", "Bento box", "Soup", "Tacos"],
        "Dinner": ["Sheet-pan chx", "Taco night", "Spaghetti", "Salmon+rice", "Pizza", "Stir-fry", "Pot roast"],
        "Snack": ["Apples", "Granola", "Crackers", "Veg & dip", "Popcorn", "Fruit", "Trail mix"],
        "Dessert": ["—", "Ice cream", "—", "Cookies", "Brownies", "Sorbet", "—"],
    }
    tx0, tx1 = x0 + pad, x1 - pad
    ty = y0 + 104
    label_w = (tx1 - tx0) * 0.13
    cellw = (tx1 - tx0 - label_w) / 7
    hdr_h = 44
    # day header
    for j, day in enumerate(days):
        hx = tx0 + label_w + j * cellw
        grad_round(img, (hx + 3, ty, hx + cellw - 3, ty + hdr_h), 8, PRIMARY_LT, PRIMARY_DK)
        d.text((hx + cellw / 2, ty + hdr_h / 2), day, font=fs(18), fill=GOLD_HI, anchor="mm")
    rh = (y1 - pad - (ty + hdr_h)) / len(meals)
    for i, meal in enumerate(meals):
        ry = ty + hdr_h + i * rh
        d.rounded_rectangle((tx0, ry + 3, tx0 + label_w - 6, ry + rh - 3), radius=8, fill=SOFT if False else (250, 245, 236), outline=GRID, width=1)
        d.text((tx0 + 12, ry + rh / 2), meal, font=fs(18), fill=ACCENT, anchor="lm")
        for j in range(7):
            hx = tx0 + label_w + j * cellw
            d.rounded_rectangle((hx + 3, ry + 3, hx + cellw - 3, ry + rh - 3), radius=8,
                                fill=WHITE if i % 2 == 0 else ROW_ALT, outline=GRID, width=1)
            val = menu[meal][j]
            col = TEXT_MUTED if val == "—" else PRIMARY
            d.text((hx + cellw / 2, ry + rh / 2), val, font=fs(17, bold=False), fill=col, anchor="mm")


SOFT = (250, 245, 236)


def content_cleaning(img, cbox):
    x0, y0, x1, y1 = cbox
    pad = 30
    d = ImageDraw.Draw(img)
    d.rectangle(cbox, fill=BG)
    d.text((x0 + pad, y0 + 22), "Cleaning & Chores", font=fs(32), fill=PRIMARY, anchor="lt")
    d.text((x0 + pad, y0 + 62), "Daily → deep cleaning + chores by family member", font=fs(19, bold=False), fill=TEXT_MUTED, anchor="lt")
    rows = [
        ("Kitchen reset", "Mom", "Daily", "Done"), ("Make beds", "All", "Daily", "Done"),
        ("Wipe counters", "Dad", "Daily", "Done"), ("Quick tidy living rm", "Kids", "Daily", "In Progress"),
        ("Vacuum main floor", "Dad", "Weekly", "Done"), ("Change bed sheets", "All", "Weekly", "Done"),
        ("Bathrooms deep wipe", "Mom", "Weekly", "Not Started"), ("Mop floors", "Dad", "Weekly", "Not Started"),
        ("Dust surfaces", "Emma", "Weekly", "Not Started"), ("Clean fridge", "Mom", "Monthly", "Not Started"),
        ("Wash windows", "Dad", "Seasonal", "Not Started"), ("Carpet shampoo", "Hired", "Deep", "Not Started"),
    ]
    headers = ["TASK", "WHO", "FREQUENCY", "STATUS"]
    tx0, tx1 = x0 + pad, x1 - pad
    ty = y0 + 104
    colx = [tx0, tx0 + (tx1 - tx0) * 0.44, tx0 + (tx1 - tx0) * 0.60, tx0 + (tx1 - tx0) * 0.82]
    hdr_h = 46
    grad_round(img, (tx0, ty, tx1, ty + hdr_h), 8, PRIMARY_LT, PRIMARY_DK)
    for i, h in enumerate(headers):
        anc = "lm" if i == 0 else "mm"
        hx = colx[i] + (14 if i == 0 else (tx1 - tx0) * 0.09)
        d.text((hx, ty + hdr_h / 2), h, font=fs(17), fill=WHITE, anchor=anc)
    rh = (y1 - pad - (ty + hdr_h)) / len(rows)
    smap = {"Done": (MINT_BG, PRIMARY), "In Progress": (WARN_BG, ACCENT), "Not Started": ((240, 236, 228), TEXT_MUTED)}
    for i, (task, who, freq, st) in enumerate(rows):
        ry = ty + hdr_h + i * rh
        if i % 2:
            d.rectangle((tx0, ry, tx1, ry + rh), fill=ROW_ALT)
        d.text((tx0 + 14, ry + rh / 2), task, font=fs(19, bold=False), fill=TEXT, anchor="lm")
        d.text((colx[1] + (tx1 - tx0) * 0.09, ry + rh / 2), who, font=fs(18, bold=False), fill=TEXT, anchor="mm")
        d.text((colx[2] + (tx1 - tx0) * 0.09, ry + rh / 2), freq, font=fs(18, bold=False), fill=TEXT_MUTED, anchor="mm")
        bg, fg = smap[st]
        cxp = colx[3] + (tx1 - tx0) * 0.09
        d.rounded_rectangle((cxp - 92, ry + rh / 2 - 17, cxp + 92, ry + rh / 2 + 17), radius=16, fill=bg)
        d.text((cxp, ry + rh / 2), st, font=fs(17), fill=fg, anchor="mm")


# ---------- the 6 renders ----------

def render_hero(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=600)
    d = ImageDraw.Draw(img)
    house_crest(img, SIZE // 2, 138, r=56)
    pill(img, SIZE // 2, 262, "THE ULTIMATE HOUSEHOLD MANAGEMENT SYSTEM", font=fs(30), pad_x=46, pad_y=20)
    wordmark(img, SIZE // 2, 402, "HOME COMMAND CENTER", 122, max_w=1780)
    gold_divider(img, SIZE // 2, 492, width=520)
    tc(d, (SIZE // 2, 540), "Finances · schedules · meals · cleaning · maintenance — one elegant dashboard",
       fs(26, bold=False), (224, 213, 190))
    chips = [("28", "POWERFUL TABS"), ("10", "LIVE DASHBOARD KPIs"), ("2-in-1", "EXCEL + SHEETS")]
    cw = 420
    total = len(chips) * cw + (len(chips) - 1) * 32
    startx = (SIZE - total) // 2 + cw // 2
    for i, (b, s) in enumerate(chips):
        stat_chip(img, startx + i * (cw + 32), 668, b, s, w=cw)
    app_window(img, (70, 770, SIZE - 70, 1888), 0, content_dashboard)
    pill(img, SIZE // 2, SIZE - 56, "28 SHEETS · INSTANT DOWNLOAD · EXCEL + GOOGLE SHEETS",
         font=fs(34), pad_x=52, pad_y=24, star=True, grad=(PRIMARY_LT, PRIMARY_DK))
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_inside(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=400)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 120, "EVERYTHING INSIDE", font=fs(38), pad_x=54, pad_y=22)
    tc(d, (SIZE // 2, 238), "28 Powerful, Connected Tabs", fserif(60), WHITE)
    gold_divider(img, SIZE // 2, 308, width=520)
    tc(d, (SIZE // 2, 352), "Not a printable — a complete household operating system that runs itself",
       fs(26, bold=False), (226, 214, 190))
    cards = [
        ("Executive Dashboard", "10 KPIs + 6 live charts"),
        ("Family Directory", "people, providers, pets"),
        ("Master Calendar", "auto week & month views"),
        ("Daily Command", "routines, top 3, wins"),
        ("Household Budget", "plan vs actual + cash flow"),
        ("Bill Center", "auto overdue alerts"),
        ("Grocery Planner", "by store, est vs actual"),
        ("Pantry Inventory", "low-stock alerts"),
        ("Fridge & Freezer", "use-by tracking"),
        ("Meal Planner", "weekly menu builder"),
        ("Recipe Library", "your family cookbook"),
        ("Cleaning Center", "daily → deep clean"),
        ("Chore Manager", "by family member"),
        ("Home Maintenance", "seasonal reminders"),
        ("Home Inventory", "for insurance"),
        ("Subscriptions", "find money leaks"),
        ("Family Goals", "progress bars"),
        ("Savings Planner", "fund-by-fund"),
        ("Holidays & Events", "gifts + budgets"),
        ("Travel Planner", "packing → itinerary"),
        ("Children's Hub", "school & growth"),
        ("Pet Care", "vet & vaccines"),
        ("Wellness", "sleep, water, habits"),
        ("Home Projects", "before/after photos"),
        ("Document Vault", "where it all lives"),
        ("Analytics", "household health score"),
        ("Inspiration Board", "paste photo ideas"),
        ("Settings", "make it yours"),
    ]
    cols = 4
    margin = 80
    gx, gy = 22, 22
    cw = (SIZE - 2 * margin - (cols - 1) * gx) // cols
    top = 440
    ch = (SIZE - top - 70 - 6 * gy) // 7
    for i, (title, sub) in enumerate(cards):
        r, ccol = divmod(i, cols)
        x = margin + ccol * (cw + gx)
        y = top + r * (ch + gy)
        shadow(img, (x, y, x + cw, y + ch), 14, 14, 45, 10)
        ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
        od = ImageDraw.Draw(ov)
        od.rounded_rectangle((x, y, x + cw, y + ch), radius=14, fill=WHITE, outline=(232, 224, 208), width=2)
        od.rounded_rectangle((x, y, x + 7, y + ch), radius=3, fill=GOLD_LT)
        od.rectangle((x + 3, y, x + 7, y + ch), fill=GOLD_LT)
        # number badge
        bx, by = x + 44, y + ch // 2
        od.ellipse((bx - 26, by - 26, bx + 26, by + 26), fill=PRIMARY)
        od.text((bx, by), str(i + 1), font=fs(24), fill=GOLD_HI, anchor="mm")
        od.text((x + 86, by - 18), title, font=fs(23), fill=PRIMARY, anchor="lm")
        od.text((x + 86, by + 18), sub, font=fs(18, bold=False), fill=TEXT_MUTED, anchor="lm")
        img.alpha_composite(ov)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_dashboard(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=360)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 116, "EXECUTIVE HOME DASHBOARD", font=fs(36), pad_x=52, pad_y=22)
    tc(d, (SIZE // 2, 232), "Your Whole Home, At A Glance", fserif(56), WHITE)
    tc(d, (SIZE // 2, 300), "10 live KPIs + spending, bills, cleaning & chore charts — all auto-updating",
       fs(25, bold=False), (226, 214, 190))
    app_window(img, (70, 400, SIZE - 70, SIZE - 70), 0, content_dashboard)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_budget_bills(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=300)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 110, "BUDGET & BILLS", font=fs(36), pad_x=52, pad_y=22)
    tc(d, (SIZE // 2, 224), "Take Control Of The Money", fserif(50), WHITE)
    # two stacked windows
    app_window(img, (60, 330, SIZE - 60, 1180), 4, content_budget)
    app_window(img, (60, 1210, SIZE - 60, SIZE - 60), 5, content_bills)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_meals_cleaning(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=300)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 110, "MEALS · CLEANING · CHORES", font=fs(34), pad_x=50, pad_y=22)
    tc(d, (SIZE // 2, 224), "Run The Home On A System", fserif(50), WHITE)
    app_window(img, (60, 330, SIZE - 60, 1140), 9, content_meals)
    app_window(img, (60, 1170, SIZE - 60, SIZE - 60), 11, content_cleaning)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_mobile(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=400)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 130, "WORKS EVERYWHERE", font=fs(38), pad_x=54, pad_y=22)
    tc(d, (SIZE // 2, 250), "Excel · Google Sheets · Mobile", fserif(56), WHITE)
    tc(d, (SIZE // 2, 320), "Update from the kitchen, the car, or the couch — your home in your pocket",
       fs(25, bold=False), (226, 214, 190))
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
    cards = [("BUDGET LEFT", "$223", PRIMARY), ("HOME HEALTH", "62%", PRIMARY),
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
    for lab, state in [("Pay water bill", True), ("Meal-prep dinner", True),
                       ("Soccer pickup 4pm", False), ("Restock pantry (6 low)", False)]:
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
        ("02_inside.png", render_inside),
        ("03_dashboard.png", render_dashboard),
        ("04_budget_bills.png", render_budget_bills),
        ("05_meals_cleaning.png", render_meals_cleaning),
        ("06_mobile.png", render_mobile),
    ]
    for name, fn in targets:
        fn(os.path.join(out_dir, name))
        print(f"  ✓ {name}")
    print(f"Wrote {len(targets)} images to {out_dir}")


if __name__ == "__main__":
    main()

"""Marketing image set for Baseball Family Command Center™ (6 images, 2000x2000).

Dense "app screenshots" that mirror the real workbook: a left sidebar of all
24 tabs, the REAL computed KPI numbers from the sample data, and fully
populated tables/charts (incl. the auto-calculated batting slash line).

  01_hero.png            - branded hero + live dashboard window
  02_inside.png          - "everything inside — 24 powerful tabs"
  03_dashboard.png       - full executive baseball dashboard
  04_budget.png          - full season budget table
  05_stats_schedule.png  - player statistics + season calendar
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
BG = (251, 248, 242)
BG_TOP = (253, 250, 246)
BG_BOT = (242, 235, 223)
WHITE = (255, 255, 255)
TEXT = (51, 51, 51)
TEXT_MUTED = (132, 126, 116)
DANGER = (201, 76, 76)
STITCH = (198, 74, 74)
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

TABS = ["Dashboard", "Player Profile", "Calendar", "Game Day", "Practice",
        "Budget", "Equipment", "Bats", "Gloves", "Stats", "Pitching",
        "Tournaments", "Travel", "Roster", "Volunteers", "Development",
        "Nutrition", "Medical", "Packing", "Fundraising", "Gallery", "Goals",
        "Analytics", "Settings"]


def fs(s, bold=True):
    return ImageFont.truetype(SANS_B if bold else SANS_R, s)


def fserif(s):
    return ImageFont.truetype(SERIF_B, s)


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


def shadow(c, box, radius, blur=24, alpha=70, dy=18):
    layer = Image.new("RGBA", c.size, (0, 0, 0, 0))
    ImageDraw.Draw(layer).rounded_rectangle(box, radius=radius, fill=(18, 50, 45, alpha))
    c.alpha_composite(layer.filter(ImageFilter.GaussianBlur(blur)), (0, dy))


def tc(d, xy, t, f, fill, anchor="mm"):
    d.text(xy, t, font=f, fill=fill, anchor=anchor)


def wordmark(c, cx, cy, text, size, max_w=None):
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


def baseball_crest(c, cx, cy, r=60, glow=True):
    if glow:
        radial_glow(c, cx, cy, int(r * 2.1), GOLD_HI, 90)
    grad_round(c, (cx - r, cy - r, cx + r, cy + r), 22, PRIMARY_LT, PRIMARY_DK, outline=GOLD_LT, width=4)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    d.rounded_rectangle((cx - r + 10, cy - r + 10, cx + r - 10, cy + r - 10), radius=16, outline=GOLD_HI, width=2)
    # two crossed bats (X) — thin handle, thick barrel
    def bat(sign):
        handle = (cx + sign * r * 0.46, cy + r * 0.54)
        mid = (cx + sign * r * 0.10, cy + r * 0.10)
        barrel = (cx - sign * r * 0.44, cy - r * 0.5)
        d.line([handle, mid], fill=GOLD_HI, width=7)
        d.line([mid, barrel], fill=GOLD_HI, width=14)
        d.ellipse((handle[0] - 6, handle[1] - 6, handle[0] + 6, handle[1] + 6), fill=GOLD_HI)
    bat(+1)
    bat(-1)
    # baseball at center
    br = r * 0.24
    d.ellipse((cx - br, cy - br, cx + br, cy + br), fill=(250, 250, 246), outline=(210, 205, 195), width=2)
    d.arc((cx - br * 0.5, cy - br, cx + br * 1.6, cy + br), 110, 250, fill=STITCH, width=3)
    d.arc((cx - br * 1.6, cy - br, cx + br * 0.5, cy + br), -70, 70, fill=STITCH, width=3)
    c.alpha_composite(ov)


def stat_chip(c, cx, cy, big, small, w=400, h=150):
    box = (cx - w // 2, cy - h // 2, cx + w // 2, cy + h // 2)
    shadow(c, box, 20, 24, 80, 16)
    grad_round(c, box, 20, PRIMARY_LT, PRIMARY_DK, outline=GOLD_LT, width=3)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    d.rounded_rectangle((box[0] + 18, box[1] + 12, box[2] - 18, box[1] + 18), radius=3, fill=GOLD_HI)
    d.text((cx, cy - h * 0.16), big, font=fserif(48), fill=GOLD_HI, anchor="mm")
    d.text((cx, cy + h * 0.28), small, font=fs(21), fill=WHITE, anchor="mm")
    c.alpha_composite(ov)


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
        h = (val / maxv) * (base - y0 - 12)
        c2 = col or (PRIMARY_LT, PRIMARY_DK)
        grad_round(img, (cx - bw * 0.28, base - h, cx + bw * 0.28, base), 8, c2[0], c2[1])
        lab_txt = f"{int(val*100)}%" if fmt_pct else str(int(val))
        d.text((cx, base - h - 20), lab_txt, font=fs(18), fill=PRIMARY, anchor="mm")
        d.text((cx, base + 18), lab, font=fs(16, bold=False), fill=TEXT_MUTED, anchor="mm")


def app_window(img, box, active_idx, content_fn, file_label="Baseball_Family_Command_Center.xlsx — Cody Reyes · Bandits 12U Select"):
    x0, y0, x1, y1 = box
    shadow(img, box, 26, 40, 95, 22)
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    od.rounded_rectangle(box, radius=24, fill=WHITE, outline=(210, 203, 190), width=2)
    img.alpha_composite(ov)
    tb_h = 58
    grad_round(img, (x0, y0, x1, y0 + tb_h + 24), 24, (54, 56, 60), (44, 46, 50))
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    od.rectangle((x0, y0 + tb_h, x1, y0 + tb_h + 4), fill=(36, 38, 42))
    for i, col in enumerate([(237, 106, 94), (245, 191, 79), (98, 197, 84)]):
        od.ellipse((x0 + 30 + i * 36, y0 + tb_h // 2 - 11, x0 + 52 + i * 36, y0 + tb_h // 2 + 11), fill=col)
    od.text(((x0 + x1) / 2, y0 + tb_h // 2), file_label, font=fs(20, bold=False), fill=(225, 222, 215), anchor="mm")
    img.alpha_composite(ov)
    sb_w = int((x1 - x0) * 0.205)
    sb = (x0, y0 + tb_h, x0 + sb_w, y1)
    grad_round(img, (sb[0], sb[1], sb[2] + 24, sb[3]), 0, PRIMARY_LT, PRIMARY_DK)
    grad_round(img, (sb[0], y1 - 24, sb[2], y1), 24, PRIMARY_DK, PRIMARY_DK)
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    bx = sb[0] + 26
    od.text((bx, sb[1] + 30), "BASEBALL FAMILY", font=fs(19), fill=GOLD_HI, anchor="lt")
    od.text((bx, sb[1] + 56), "24-tab system", font=fs(15, bold=False), fill=(170, 200, 192), anchor="lt")
    od.line((sb[0] + 20, sb[1] + 88, sb[2] - 16, sb[1] + 88), fill=(255, 255, 255, 40), width=1)
    list_top = sb[1] + 102
    rowh = (y1 - 24 - list_top) / len(TABS)
    palette = [HIGHLIGHT, GOLD_HI, SURFACE, (150, 200, 190)]
    for i, name in enumerate(TABS):
        ry = list_top + i * rowh
        if i == active_idx:
            od.rounded_rectangle((sb[0] + 12, ry + 2, sb[2] - 10, ry + rowh - 2), radius=8, fill=(255, 255, 255, 235))
            od.rounded_rectangle((sb[0] + 12, ry + 2, sb[0] + 19, ry + rowh - 2), radius=3, fill=GOLD_HI)
            dotc = PRIMARY; txtc = PRIMARY; font = fs(18)
        else:
            dotc = palette[i % len(palette)]; txtc = (214, 226, 222); font = fs(17, bold=False)
        cyr = ry + rowh / 2
        od.ellipse((sb[0] + 30, cyr - 6, sb[0] + 42, cyr + 6), fill=dotc)
        od.text((sb[0] + 56, cyr), name, font=font, fill=txtc, anchor="lm")
    img.alpha_composite(ov)
    cbox = (sb[2] + 1, y0 + tb_h + 4, x1, y1)
    content_fn(img, cbox)


KPIS = [
    ("DAYS TO NEXT GAME", "2", "vs Bandits"),
    ("GAMES / WEEK", "2", "this week"),
    ("PRACTICES / WEEK", "3", "this week"),
    ("NEXT TOURNAMENT", "14", "Summer Slam"),
    ("MONTHLY BUDGET", "$900", "set in Settings"),
    ("BUDGET LEFT", "$2,455", "of $7,410"),
    ("EQUIP. ALERTS", "3", "replace soon"),
    ("GAMES PLAYED", "4", "of 8"),
    ("TRAVEL MILES", "485", "this season"),
    ("VOLUNTEER HRS", "18", "logged"),
    ("DEV PROGRESS", "63%", "skills avg"),
    ("SEASON COMPLETE", "50%", "4 / 8 games"),
]


def content_dashboard(img, cbox):
    x0, y0, x1, y1 = cbox
    pad = 30
    d = ImageDraw.Draw(img)
    d.rectangle(cbox, fill=BG)
    d.text((x0 + pad, y0 + 22), "Executive Baseball Dashboard", font=fs(33), fill=PRIMARY, anchor="lt")
    d.text((x0 + pad, y0 + 64), "Cody Reyes · Bandits 12U Select  ·  live KPIs update as you type", font=fs(19, bold=False), fill=TEXT_MUTED, anchor="lt")
    d.rounded_rectangle((x1 - pad - 150, y0 + 26, x1 - pad, y0 + 62), radius=18, fill=MINT_BG)
    d.text((x1 - pad - 75, y0 + 44), "● live", font=fs(18), fill=PRIMARY, anchor="mm")
    gx = x0 + pad; gy = y0 + 98
    gw = (x1 - x0 - 2 * pad); gap = 14
    kw = (gw - 5 * gap) / 6; kh = 116
    for i, (lab, val, sub) in enumerate(KPIS):
        r, ci = divmod(i, 6)
        kx = gx + ci * (kw + gap); ky = gy + r * (kh + gap)
        d.rounded_rectangle((kx, ky, kx + kw, ky + kh), radius=12, fill=WHITE, outline=GRID, width=2)
        d.rounded_rectangle((kx + 12, ky, kx + kw - 12, ky + 5), radius=2, fill=GOLD_LT)
        d.text((kx + 14, ky + 16), lab, font=fs(12), fill=ACCENT, anchor="lt")
        d.text((kx + 14, ky + 58), val, font=fserif(30), fill=PRIMARY, anchor="lm")
        d.text((kx + 14, ky + 96), sub, font=fs(12, bold=False), fill=TEXT_MUTED, anchor="lm")
    cy_top = gy + 2 * (kh + gap) + 18
    d.text((gx, cy_top), "MONEY · SCHEDULE · DEVELOPMENT", font=fs(20), fill=ACCENT, anchor="lt")
    panels_y = cy_top + 34
    panel_h = (y1 - panels_y - pad)
    pw = (gw - 3 * gap) / 4
    px = gx
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Baseball Spending", font=fs(17), fill=ACCENT, anchor="lt")
    donut(d, px + pw * 0.42, panels_y + panel_h * 0.52, min(panel_h * 0.28, pw * 0.27),
          [(35, PRIMARY), (29, ACCENT), (17, HIGHLIGHT), (13, SURFACE), (5, (170, 150, 120))], "$5.0K", "spent")
    legend(d, px + pw * 0.04, panels_y + panel_h - 104, [(PRIMARY, "Fees 35%"), (ACCENT, "Travel 29%"), (HIGHLIGHT, "Gear 17%")], 15, 30)
    px = gx + (pw + gap)
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Games vs Practices", font=fs(17), fill=ACCENT, anchor="lt")
    donut(d, px + pw * 0.42, panels_y + panel_h * 0.52, min(panel_h * 0.28, pw * 0.27),
          [(8, PRIMARY), (3, HIGHLIGHT), (1, GOLD_LT), (6, SURFACE)], "18", "events")
    legend(d, px + pw * 0.04, panels_y + panel_h - 104, [(PRIMARY, "Games 8"), (HIGHLIGHT, "Practice 3"), (SURFACE, "Other 6")], 15, 30)
    px = gx + 2 * (pw + gap)
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Skill Progress", font=fs(17), fill=ACCENT, anchor="lt")
    vbars(img, d, (px + 18, panels_y + 50, px + pw - 12, panels_y + panel_h - 6),
          [("Hit", 0.72, (HIGHLIGHT, (70, 200, 165))), ("Field", 0.70, (HIGHLIGHT, (70, 200, 165))),
           ("Throw", 0.58, (HIGHLIGHT, (70, 200, 165))), ("Run", 0.75, (HIGHLIGHT, (70, 200, 165))),
           ("Speed", 0.55, (HIGHLIGHT, (70, 200, 165))), ("Str", 0.45, (HIGHLIGHT, (70, 200, 165)))])
    px = gx + 3 * (pw + gap)
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Batting Totals", font=fs(17), fill=ACCENT, anchor="lt")
    vbars(img, d, (px + 24, panels_y + 50, px + pw - 16, panels_y + panel_h - 6),
          [("Hits", 6, None), ("RBI", 5, None), ("Runs", 4, None), ("SB", 2, None)], maxv=6, fmt_pct=False)


def content_budget(img, cbox):
    x0, y0, x1, y1 = cbox
    pad = 30
    d = ImageDraw.Draw(img)
    d.rectangle(cbox, fill=BG)
    d.text((x0 + pad, y0 + 20), "Baseball Budget", font=fs(32), fill=PRIMARY, anchor="lt")
    d.text((x0 + pad, y0 + 60), "Every season cost in one place  ·  cost-per-game & remaining, automatic", font=fs(18, bold=False), fill=TEXT_MUTED, anchor="lt")
    strip = [("SEASON BUDGET", "$7,410"), ("SPENT", "$4,955"), ("REMAINING", "$2,455"), ("COST / GAME", "$619")]
    sy = y0 + 94
    sw = (x1 - x0 - 2 * pad - 3 * 14) / 4
    for i, (lab, val) in enumerate(strip):
        sx = x0 + pad + i * (sw + 14)
        d.rounded_rectangle((sx, sy, sx + sw, sy + 92), radius=10, fill=WHITE, outline=GRID, width=2)
        d.rounded_rectangle((sx + 12, sy, sx + sw - 12, sy + 5), radius=2, fill=GOLD_LT)
        d.text((sx + sw / 2, sy + 30), lab, font=fs(15), fill=ACCENT, anchor="mm")
        d.text((sx + sw / 2, sy + 64), val, font=fserif(28), fill=PRIMARY, anchor="mm")
    rows = [
        ("Registration", 250, 250), ("League Fees", 600, 600), ("Uniforms", 180, 180),
        ("Cleats", 90, 90), ("Gloves", 200, 200), ("Bats", 350, 200), ("Helmets", 60, 60),
        ("Batting Gloves", 60, 35), ("Bags", 80, 80), ("Protective Gear", 120, 0),
        ("Tournaments", 1400, 900), ("Hotels", 1200, 720), ("Fuel", 600, 380),
        ("Meals", 700, 360), ("Private Lessons", 800, 480), ("Batting Cage", 300, 180),
        ("Team Apparel", 150, 150), ("Photos", 120, 0), ("Fundraising", 0, 0),
        ("Miscellaneous", 150, 90),
    ]
    ty = sy + 116
    tx0, tx1 = x0 + pad, x1 - pad
    headers = ["CATEGORY", "BUDGET", "SPENT", "REMAINING", "% USED"]
    colx = [tx0, tx0 + (tx1 - tx0) * 0.42, tx0 + (tx1 - tx0) * 0.56, tx0 + (tx1 - tx0) * 0.70, tx0 + (tx1 - tx0) * 0.84]
    hdr_h = 42
    grad_round(img, (tx0, ty, tx1, ty + hdr_h), 8, PRIMARY_LT, PRIMARY_DK)
    for i, h in enumerate(headers):
        anc = "lm" if i == 0 else "mm"
        hx = colx[i] + (14 if i == 0 else (tx1 - tx0) * 0.07)
        d.text((hx, ty + hdr_h / 2), h, font=fs(16), fill=WHITE, anchor=anc)
    rh = (y1 - pad - (ty + hdr_h)) / (len(rows) + 1)
    barw = (tx1 - tx0) * 0.11
    for i, (cat, pl, sp) in enumerate(rows):
        ry = ty + hdr_h + i * rh
        if i % 2:
            d.rectangle((tx0, ry, tx1, ry + rh), fill=ROW_ALT)
        d.text((tx0 + 14, ry + rh / 2), cat, font=fs(16, bold=False), fill=TEXT, anchor="lm")
        d.text((colx[1] + (tx1 - tx0) * 0.07, ry + rh / 2), f"${pl:,}", font=fs(16, bold=False), fill=TEXT, anchor="mm")
        d.text((colx[2] + (tx1 - tx0) * 0.07, ry + rh / 2), f"${sp:,}", font=fs(16), fill=PRIMARY, anchor="mm")
        d.text((colx[3] + (tx1 - tx0) * 0.07, ry + rh / 2), f"${pl-sp:,}", font=fs(16, bold=False), fill=TEXT, anchor="mm")
        used = sp / pl if pl else 0
        tbx0 = colx[4] + 4
        d.rounded_rectangle((tbx0, ry + rh / 2 - 9, tbx0 + barw, ry + rh / 2 + 9), radius=8, fill=(236, 230, 220))
        grad_round(img, (tbx0, ry + rh / 2 - 9, tbx0 + barw * min(used, 1), ry + rh / 2 + 9), 8, PRIMARY_LT, PRIMARY_DK)
        d.text((tbx0 + barw + 12, ry + rh / 2), f"{int(used*100)}%", font=fs(14), fill=ACCENT, anchor="lm")
    ry = ty + hdr_h + len(rows) * rh
    d.rectangle((tx0, ry, tx1, ry + rh), fill=SURFACE)
    d.text((tx0 + 14, ry + rh / 2), "TOTAL", font=fs(18), fill=PRIMARY, anchor="lm")
    d.text((colx[1] + (tx1 - tx0) * 0.07, ry + rh / 2), "$7,410", font=fs(17), fill=PRIMARY, anchor="mm")
    d.text((colx[2] + (tx1 - tx0) * 0.07, ry + rh / 2), "$4,955", font=fs(17), fill=PRIMARY, anchor="mm")
    d.text((colx[3] + (tx1 - tx0) * 0.07, ry + rh / 2), "$2,455", font=fs(17), fill=PRIMARY, anchor="mm")
    d.text((colx[4] + 4 + barw + 12, ry + rh / 2), "67%", font=fs(15), fill=PRIMARY, anchor="lm")


def content_stats(img, cbox):
    x0, y0, x1, y1 = cbox
    pad = 30
    d = ImageDraw.Draw(img)
    d.rectangle(cbox, fill=BG)
    d.text((x0 + pad, y0 + 20), "Player Statistics", font=fs(32), fill=PRIMARY, anchor="lt")
    d.text((x0 + pad, y0 + 60), "Log every game — AVG, OBP, SLG & OPS calculate themselves", font=fs(18, bold=False), fill=TEXT_MUTED, anchor="lt")
    # slash-line cards
    slash = [("AVG", ".429"), ("OBP", ".529"), ("SLG", ".786"), ("OPS", "1.315")]
    sy = y0 + 94
    sw = (x1 - x0 - 2 * pad - 3 * 14) / 4
    for i, (lab, val) in enumerate(slash):
        sx = x0 + pad + i * (sw + 14)
        grad_round(img, (sx, sy, sx + sw, sy + 96), 12, PRIMARY_LT, PRIMARY_DK, outline=GOLD_LT, width=2)
        d.rounded_rectangle((sx + 14, sy + 10, sx + sw - 14, sy + 16), radius=3, fill=GOLD_HI)
        d.text((sx + sw / 2, sy + 40), lab, font=fs(17), fill=GOLD_HI, anchor="mm")
        d.text((sx + sw / 2, sy + 72), val, font=fserif(30), fill=WHITE, anchor="mm")
    # game log
    rows = [
        ("vs Storm", 4, 2, 1, 1, 0, 0, 1, 1),
        ("vs Cobras", 3, 1, 1, 0, 0, 1, 0, 0),
        ("vs Rockets", 4, 2, 2, 3, 1, 0, 1, 0),
        ("vs Titans", 3, 1, 0, 1, 0, 1, 1, 1),
    ]
    headers = ["OPPONENT", "AB", "H", "R", "RBI", "HR", "BB", "K", "SB"]
    tx0, tx1 = x0 + pad, x1 - pad
    ty = sy + 120
    n = len(headers)
    colw = (tx1 - tx0) / n
    hdr_h = 44
    grad_round(img, (tx0, ty, tx1, ty + hdr_h), 8, PRIMARY_LT, PRIMARY_DK)
    for i, h in enumerate(headers):
        anc = "lm" if i == 0 else "mm"
        hx = tx0 + (18 if i == 0 else colw * (i + 0.5))
        d.text((hx, ty + hdr_h / 2), h, font=fs(16), fill=WHITE, anchor=anc)
    rh = (y1 - pad - (ty + hdr_h)) / (len(rows) + 1)
    for i, row in enumerate(rows):
        ry = ty + hdr_h + i * rh
        if i % 2:
            d.rectangle((tx0, ry, tx1, ry + rh), fill=ROW_ALT)
        d.text((tx0 + 18, ry + rh / 2), row[0], font=fs(18), fill=PRIMARY, anchor="lm")
        for ci in range(1, n):
            d.text((tx0 + colw * (ci + 0.5), ry + rh / 2), str(row[ci]), font=fs(17, bold=False), fill=TEXT, anchor="mm")
    ry = ty + hdr_h + len(rows) * rh
    d.rectangle((tx0, ry, tx1, ry + rh), fill=SURFACE)
    d.text((tx0 + 18, ry + rh / 2), "SEASON TOTALS", font=fs(16), fill=PRIMARY, anchor="lm")
    totals = [14, 6, 4, 5, 1, 3, 3, 2]
    for ci, val in enumerate(totals, 1):
        d.text((tx0 + colw * (ci + 0.5), ry + rh / 2), str(val), font=fs(17), fill=PRIMARY, anchor="mm")


def content_calendar(img, cbox):
    x0, y0, x1, y1 = cbox
    pad = 30
    d = ImageDraw.Draw(img)
    d.rectangle(cbox, fill=BG)
    d.text((x0 + pad, y0 + 22), "Master Season Calendar", font=fs(32), fill=PRIMARY, anchor="lt")
    d.text((x0 + pad, y0 + 62), "Games, practices & tournaments — countdowns & conflict flags", font=fs(18, bold=False), fill=TEXT_MUTED, anchor="lt")
    rows = [
        ("Practice", "Team practice", "in 1d", "Miller Field", ""),
        ("Game", "vs Bandits", "in 2d", "Miller Field", ""),
        ("Fielding", "Fielding focus", "in 3d", "Miller Field", ""),
        ("Batting Cage", "Cage session", "in 4d", "Grand Slam Cages", ""),
        ("Game", "vs Rockets", "in 5d", "Eastside Park", "⚠ clash"),
        ("School", "Science fair", "in 5d", "Lincoln Middle", "⚠ clash"),
        ("Practice", "Pitching + pen", "in 7d", "Miller Field", ""),
        ("Game", "vs Cobras", "in 9d", "North Diamond", ""),
        ("Game", "vs Storm", "in 12d", "Miller Field", ""),
        ("Tournament", "Summer Slam", "in 14d", "Capital City", "⚠ clash"),
    ]
    headers = ["TYPE", "EVENT", "WHEN", "LOCATION", ""]
    tx0, tx1 = x0 + pad, x1 - pad
    ty = y0 + 102
    colx = [tx0, tx0 + (tx1 - tx0) * 0.22, tx0 + (tx1 - tx0) * 0.52, tx0 + (tx1 - tx0) * 0.66, tx0 + (tx1 - tx0) * 0.92]
    hdr_h = 44
    grad_round(img, (tx0, ty, tx1, ty + hdr_h), 8, PRIMARY_LT, PRIMARY_DK)
    for i, h in enumerate(headers):
        d.text((colx[i] + (16 if i in (0, 1, 3) else 0), ty + hdr_h / 2), h, font=fs(16), fill=WHITE, anchor="lm" if i in (0, 1, 3) else "mm")
    rh = (y1 - pad - (ty + hdr_h)) / len(rows)
    tmap = {"Game": (PRIMARY, WHITE), "Practice": (HIGHLIGHT, PRIMARY), "Tournament": (GOLD_LT, PRIMARY),
            "Fielding": (MINT_BG, PRIMARY), "Batting Cage": (MINT_BG, PRIMARY), "School": ((235, 230, 222), TEXT_MUTED)}
    for i, (typ, ev, when, loc, flag) in enumerate(rows):
        ry = ty + hdr_h + i * rh
        if flag:
            d.rectangle((tx0, ry, tx1, ry + rh), fill=WARN_BG)
        elif i % 2:
            d.rectangle((tx0, ry, tx1, ry + rh), fill=ROW_ALT)
        bg, fg = tmap.get(typ, ((235, 230, 222), TEXT_MUTED))
        d.rounded_rectangle((colx[0] + 12, ry + rh / 2 - 16, colx[0] + 12 + 150, ry + rh / 2 + 16), radius=15, fill=bg)
        d.text((colx[0] + 12 + 75, ry + rh / 2), typ, font=fs(15), fill=fg, anchor="mm")
        d.text((colx[1] + 16, ry + rh / 2), ev, font=fs(18), fill=PRIMARY, anchor="lm")
        d.text((colx[2], ry + rh / 2), when, font=fs(17), fill=ACCENT, anchor="mm")
        d.text((colx[3] + 16, ry + rh / 2), loc, font=fs(16, bold=False), fill=TEXT_MUTED, anchor="lm")
        if flag:
            d.text((colx[4], ry + rh / 2), flag, font=fs(15), fill=DANGER, anchor="mm")


# ---------- renders ----------

def render_hero(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=640)
    d = ImageDraw.Draw(img)
    baseball_crest(img, SIZE // 2, 132, r=56)
    pill(img, SIZE // 2, 256, "THE ULTIMATE BASEBALL SEASON OPERATING SYSTEM", font=fs(28), pad_x=44, pad_y=20)
    wordmark(img, SIZE // 2, 372, "BASEBALL FAMILY", 104, max_w=1520)
    wordmark(img, SIZE // 2, 470, "COMMAND CENTER", 104, max_w=1520)
    gold_divider(img, SIZE // 2, 540, width=520)
    tc(d, (SIZE // 2, 588), "Schedule · budget · equipment · stats · travel — one season, one system",
       fs(26, bold=False), (224, 213, 190))
    chips = [("24", "POWERFUL TABS"), ("12", "LIVE DASHBOARD KPIs"), ("2-in-1", "EXCEL + SHEETS")]
    cw = 420
    total = len(chips) * cw + (len(chips) - 1) * 32
    startx = (SIZE - total) // 2 + cw // 2
    for i, (b, s) in enumerate(chips):
        stat_chip(img, startx + i * (cw + 32), 706, b, s, w=cw)
    app_window(img, (70, 800, SIZE - 70, 1900), 0, content_dashboard)
    pill(img, SIZE // 2, SIZE - 52, "24 SHEETS · INSTANT DOWNLOAD · EXCEL + GOOGLE SHEETS",
         font=fs(33), pad_x=50, pad_y=24, star=True, grad=(PRIMARY_LT, PRIMARY_DK))
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_inside(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=400)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 120, "EVERYTHING INSIDE", font=fs(38), pad_x=54, pad_y=22)
    tc(d, (SIZE // 2, 238), "24 Powerful, Connected Tabs", fserif(58), WHITE)
    gold_divider(img, SIZE // 2, 308, width=520)
    tc(d, (SIZE // 2, 352), "Not a planner — a complete baseball-season operating system that runs itself",
       fs(25, bold=False), (226, 214, 190))
    cards = [
        ("Executive Dashboard", "12 KPIs + 4 live charts"), ("Player Profile", "stats card & medical"),
        ("Season Calendar", "countdowns + conflicts"), ("Game Day Center", "ballpark, times, score"),
        ("Practice Planner", "attendance & hours"), ("Baseball Budget", "20 cost categories"),
        ("Equipment Center", "replacement alerts"), ("Bat Inventory", "drop, cert, usage"),
        ("Glove Care Log", "conditioning & lacing"), ("Player Statistics", "AVG · OBP · SLG · OPS"),
        ("Pitching Tracker", "pitch count + ERA"), ("Tournaments", "fees, hotels, schedule"),
        ("Travel Planner", "miles, fuel, lodging"), ("Team Roster", "players & parents"),
        ("Snack & Volunteer", "duties + hours"), ("Player Development", "skill progress bars"),
        ("Nutrition", "game-day fuel"), ("Medical Center", "injuries & recovery"),
        ("Packing Checklist", "auto-built lists"), ("Fundraising", "offset the cost"),
        ("Photo Gallery", "season memories"), ("Baseball Goals", "season & skill goals"),
        ("Analytics", "readiness score"), ("Settings", "make it yours"),
    ]
    cols = 4
    margin = 80
    gx, gy = 22, 22
    cw = (SIZE - 2 * margin - (cols - 1) * gx) // cols
    top = 440
    ch = (SIZE - top - 60 - 5 * gy) // 6
    for i, (title, sub) in enumerate(cards):
        r, ccol = divmod(i, cols)
        x = margin + ccol * (cw + gx); y = top + r * (ch + gy)
        shadow(img, (x, y, x + cw, y + ch), 14, 13, 42, 9)
        ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
        od = ImageDraw.Draw(ov)
        od.rounded_rectangle((x, y, x + cw, y + ch), radius=14, fill=WHITE, outline=(232, 224, 208), width=2)
        od.rounded_rectangle((x, y, x + 7, y + ch), radius=3, fill=GOLD_LT)
        od.rectangle((x + 3, y, x + 7, y + ch), fill=GOLD_LT)
        cyc = y + ch // 2
        bx = x + 42
        od.ellipse((bx - 23, cyc - 23, bx + 23, cyc + 23), fill=PRIMARY)
        od.text((bx, cyc), str(i + 1), font=fs(21), fill=GOLD_HI, anchor="mm")
        od.text((x + 82, cyc - 17), title, font=fs(21), fill=PRIMARY, anchor="lm")
        od.text((x + 82, cyc + 22), sub, font=fs(16, bold=False), fill=TEXT_MUTED, anchor="lm")
        img.alpha_composite(ov)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_dashboard(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=360)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 116, "EXECUTIVE BASEBALL DASHBOARD", font=fs(33), pad_x=50, pad_y=22)
    tc(d, (SIZE // 2, 232), "Your Whole Season, At A Glance", fserif(54), WHITE)
    tc(d, (SIZE // 2, 300), "12 live KPIs + spending, schedule, skill & batting charts — all auto-updating",
       fs(24, bold=False), (226, 214, 190))
    app_window(img, (70, 400, SIZE - 70, SIZE - 70), 0, content_dashboard)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_budget(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=320)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 112, "BASEBALL BUDGET", font=fs(36), pad_x=52, pad_y=22)
    tc(d, (SIZE // 2, 230), "Know Exactly Where It Goes", fserif(52), WHITE)
    tc(d, (SIZE // 2, 290), "All 20 season cost categories — spending, remaining & cost-per-game",
       fs(23, bold=False), (226, 214, 190))
    app_window(img, (80, 360, SIZE - 80, SIZE - 70), 5, content_budget)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_stats_schedule(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=300)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 110, "STATS & SCHEDULE", font=fs(36), pad_x=52, pad_y=22)
    tc(d, (SIZE // 2, 224), "Track Every Hit & Never Miss A Game", fserif(46), WHITE)
    app_window(img, (60, 330, SIZE - 60, 1150), 9, content_stats)
    app_window(img, (60, 1180, SIZE - 60, SIZE - 60), 2, content_calendar)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_mobile(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=400)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 130, "WORKS EVERYWHERE", font=fs(38), pad_x=54, pad_y=22)
    tc(d, (SIZE // 2, 250), "Excel · Google Sheets · Mobile", fserif(56), WHITE)
    tc(d, (SIZE // 2, 320), "Check the schedule from the dugout, the car, or the stands — season in your pocket",
       fs(24, bold=False), (226, 214, 190))
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
    od.text(((sx0 + sx1) // 2, sy0 + 56), "Baseball Command", font=fserif(32), fill=GOLD_HI, anchor="mm")
    y = sy0 + 150
    cards = [("DAYS TO NEXT GAME", "2", PRIMARY), ("BATTING AVG", ".429", PRIMARY),
             ("BUDGET LEFT", "$2,455", ACCENT), ("SEASON COMPLETE", "50%", PRIMARY)]
    for lab, val, col in cards:
        cb = (sx0 + 30, y, sx1 - 30, y + 135)
        od.rounded_rectangle(cb, radius=18, fill=WHITE, outline=GOLD_LT, width=2)
        od.rounded_rectangle((cb[0] + 20, y, cb[2] - 20, y + 5), radius=2, fill=GOLD_LT)
        od.text((cb[0] + 26, y + 34), lab, font=fs(22), fill=ACCENT, anchor="lt")
        od.text((cb[0] + 26, y + 94), val, font=fserif(46), fill=col, anchor="lm")
        y += 152
    od.text((sx0 + 40, y + 16), "THIS WEEK", font=fs(22), fill=ACCENT, anchor="lt")
    y += 52
    for lab, state in [("Snack duty — vs Bandits", False), ("Pack game bag", False),
                       ("Batting cage 7pm", True), ("Pay Summer Slam fee", True)]:
        col = HIGHLIGHT if state else BG
        od.ellipse((sx0 + 40, y + 6, sx0 + 78, y + 44), fill=col, outline=PRIMARY, width=3)
        if state:
            od.text((sx0 + 59, y + 24), "✓", font=fs(24), fill=PRIMARY, anchor="mm")
        od.text((sx0 + 96, y + 24), lab, font=fs(22, bold=False), fill=TEXT, anchor="lm")
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
        ("04_budget.png", render_budget),
        ("05_stats_schedule.png", render_stats_schedule),
        ("06_mobile.png", render_mobile),
    ]
    for name, fn in targets:
        fn(os.path.join(out_dir, name))
        print(f"  ✓ {name}")
    print(f"Wrote {len(targets)} images to {out_dir}")


if __name__ == "__main__":
    main()

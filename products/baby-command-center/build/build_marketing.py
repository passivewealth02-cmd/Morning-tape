"""Marketing image set for Baby Command Center™ (6 images, 2000x2000).

Dense app-screenshot marketing that mirrors the real workbook: a left sidebar
of all 21 tabs, the REAL computed KPI numbers from the sample data, and fully
populated tables/charts (feeding, sleep, growth, milestones).

  01_hero.png            - branded hero + live baby dashboard
  02_inside.png          - "everything inside — 21 gentle tabs"
  03_dashboard.png       - full executive baby dashboard
  04_trackers.png        - feeding + diaper trackers
  05_growth_miles.png    - growth + milestones
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

TABS = ["Dashboard", "Profile", "Daily Log", "Feeding", "Sleep", "Diapers",
        "Growth", "Milestones", "Medical", "Appointments", "Budget", "Shopping",
        "Inventory", "Childcare", "Routine", "Development", "Memories", "Travel",
        "Family Goals", "Analytics", "Settings"]


def fs(s, bold=True):
    return ImageFont.truetype(SANS_B if bold else SANS_R, s)


def fserif(s):
    return ImageFont.truetype(SERIF_B, s)


def vgradient(w, h, top, bottom):
    col = Image.new("RGB", (1, h)); px = col.load()
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
    layer = Image.new("RGBA", c.size, (0, 0, 0, 0)); layer.paste(g, (x0, y0), mask)
    c.alpha_composite(layer)
    if outline and width:
        ImageDraw.Draw(c).rounded_rectangle((x0, y0, x1, y1), radius=radius, outline=outline, width=width)


def radial_glow(c, cx, cy, r, color, strength=120):
    layer = Image.new("RGBA", c.size, (0, 0, 0, 0))
    ImageDraw.Draw(layer).ellipse((cx - r, cy - r, cx + r, cy + r), fill=color + (strength,))
    c.alpha_composite(layer.filter(ImageFilter.GaussianBlur(r // 2)))


def premium_bg(c, band_h=0):
    c.alpha_composite(vgradient(c.width, c.height, BG_TOP, BG_BOT))
    dots = Image.new("RGBA", c.size, (0, 0, 0, 0)); dd = ImageDraw.Draw(dots)
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
    wm = Image.new("RGBA", c.size, (0, 0, 0, 0)); wd = ImageDraw.Draw(wm)
    for rr in (300, 230, 160):
        wd.ellipse((c.width - 120 - rr, band_h - 60 - rr, c.width - 120 + rr, band_h - 60 + rr), outline=(255, 255, 255, 22), width=3)
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
    bb = d.textbbox((cx, cy), text, font=f, anchor="mm"); pad = 30
    bx0, by0, bx1, by1 = bb[0] - pad, bb[1] - pad, bb[2] + pad, bb[3] + pad
    w, h = bx1 - bx0, by1 - by0
    grad = vgradient(w, h, GOLD_HI, GOLD)
    mask = Image.new("L", (w, h), 0)
    ImageDraw.Draw(mask).text((cx - bx0, cy - by0), text, font=f, fill=255, anchor="mm")
    layer = Image.new("RGBA", c.size, (0, 0, 0, 0)); layer.paste(grad, (bx0, by0), mask)
    c.alpha_composite(layer)


def pill(c, cx, cy, text, font, pad_x=60, pad_y=26, star=False, fg=WHITE, grad=(GOLD_LT, GOLD), outline=GOLD_HI):
    d = ImageDraw.Draw(c)
    label = f"★  {text}" if star else text
    tw = d.textlength(label, font=font); th = font.size
    w, h = tw + pad_x * 2, th + pad_y * 2
    box = (cx - w // 2, cy - h // 2, cx + w // 2, cy + h // 2)
    shadow(c, box, h // 2, 22, 70, 12)
    grad_round(c, box, h // 2, grad[0], grad[1])
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
    od.rounded_rectangle((box[0] + 5, box[1] + 5, box[2] - 5, box[3] - 5), radius=(h - 10) // 2, outline=outline, width=2)
    od.text((cx, cy), label, font=font, fill=fg, anchor="mm")
    c.alpha_composite(ov)


def gold_divider(c, cx, cy, width=560, color=GOLD_HI):
    d = ImageDraw.Draw(c)
    d.line((cx - width // 2, cy, cx - 30, cy), fill=color, width=3)
    d.line((cx + 30, cy, cx + width // 2, cy), fill=color, width=3)
    d.polygon([(cx, cy - 12), (cx + 16, cy), (cx, cy + 12), (cx - 16, cy)], fill=color)


def bottle_crest(c, cx, cy, r=60, glow=True):
    if glow:
        radial_glow(c, cx, cy, int(r * 2.1), GOLD_HI, 90)
    grad_round(c, (cx - r, cy - r, cx + r, cy + r), 22, PRIMARY_LT, PRIMARY_DK, outline=GOLD_LT, width=4)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0)); d = ImageDraw.Draw(ov)
    d.rounded_rectangle((cx - r + 10, cy - r + 10, cx + r - 10, cy + r - 10), radius=16, outline=GOLD_HI, width=2)
    # baby bottle: nipple + collar + body with measurement ticks
    bw = r * 0.42
    d.rounded_rectangle((cx - bw * 0.4, cy - r * 0.62, cx + bw * 0.4, cy - r * 0.42), radius=6, fill=GOLD_HI)  # nipple
    d.rounded_rectangle((cx - bw * 0.62, cy - r * 0.44, cx + bw * 0.62, cy - r * 0.30), radius=4, fill=GOLD_HI)  # collar
    d.rounded_rectangle((cx - bw, cy - r * 0.28, cx + bw, cy + r * 0.6), radius=14, fill=HIGHLIGHT, outline=GOLD_HI, width=3)  # body
    for i in range(3):
        yy = cy - r * 0.08 + i * r * 0.18
        d.line((cx - bw * 0.55, yy, cx - bw * 0.15, yy), fill=PRIMARY, width=4)
    c.alpha_composite(ov)


def stat_chip(c, cx, cy, big, small, w=400, h=150):
    box = (cx - w // 2, cy - h // 2, cx + w // 2, cy + h // 2)
    shadow(c, box, 20, 24, 80, 16)
    grad_round(c, box, 20, PRIMARY_LT, PRIMARY_DK, outline=GOLD_LT, width=3)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0)); d = ImageDraw.Draw(ov)
    d.rounded_rectangle((box[0] + 18, box[1] + 12, box[2] - 18, box[1] + 18), radius=3, fill=GOLD_HI)
    d.text((cx, cy - h * 0.16), big, font=fserif(46), fill=GOLD_HI, anchor="mm")
    d.text((cx, cy + h * 0.28), small, font=fs(21), fill=WHITE, anchor="mm")
    c.alpha_composite(ov)


def donut(d, cx, cy, r, segs, center_top=None, center_sub=None, hole=0.55):
    ang = -90
    for pct, col in segs:
        s = pct * 3.6
        d.pieslice((cx - r, cy - r, cx + r, cy + r), ang, ang + s, fill=col); ang += s
    hr = r * hole
    d.ellipse((cx - hr, cy - hr, cx + hr, cy + hr), fill=WHITE)
    if center_top:
        d.text((cx, cy - (10 if center_sub else 0)), center_top, font=fserif(int(r * 0.32)), fill=PRIMARY, anchor="mm")
    if center_sub:
        d.text((cx, cy + int(r * 0.26)), center_sub, font=fs(int(r * 0.13)), fill=TEXT_MUTED, anchor="mm")


def legend(d, x, y, items, fsz=20, gap=40):
    for i, (col, lab) in enumerate(items):
        yy = y + i * gap
        d.rounded_rectangle((x, yy, x + 24, yy + 24), radius=5, fill=col)
        d.text((x + 36, yy + 12), lab, font=fs(fsz, bold=False), fill=TEXT, anchor="lm")


def mini_lines(d, box, series, colors):
    x0, y0, x1, y1 = box
    allv = [v for s in series for v in s]
    mn, mx = min(allv), max(allv)
    rng = (mx - mn) or 1
    n = len(series[0])
    for si, s in enumerate(series):
        pts = []
        for i, v in enumerate(s):
            px = x0 + (x1 - x0) * (i / (n - 1))
            py = y1 - (y1 - y0) * ((v - mn) / rng)
            pts.append((px, py))
        d.line(pts, fill=colors[si], width=5, joint="curve")
        for p in pts:
            d.ellipse((p[0] - 5, p[1] - 5, p[0] + 5, p[1] + 5), fill=colors[si])


def fit_font(d, text, max_w, start, serif=True):
    s = start; f = fserif(s) if serif else fs(s)
    while s > 12 and d.textlength(text, font=f) > max_w:
        s -= 1; f = fserif(s) if serif else fs(s)
    return f


def app_window(img, box, active_idx, content_fn, file_label="Baby_Command_Center.xlsx — Baby Rose · 4 months"):
    x0, y0, x1, y1 = box
    shadow(img, box, 26, 40, 95, 22)
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
    od.rounded_rectangle(box, radius=24, fill=WHITE, outline=(210, 203, 190), width=2)
    img.alpha_composite(ov)
    tb_h = 58
    grad_round(img, (x0, y0, x1, y0 + tb_h + 24), 24, (54, 56, 60), (44, 46, 50))
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
    od.rectangle((x0, y0 + tb_h, x1, y0 + tb_h + 4), fill=(36, 38, 42))
    for i, col in enumerate([(237, 106, 94), (245, 191, 79), (98, 197, 84)]):
        od.ellipse((x0 + 30 + i * 36, y0 + tb_h // 2 - 11, x0 + 52 + i * 36, y0 + tb_h // 2 + 11), fill=col)
    od.text(((x0 + x1) / 2, y0 + tb_h // 2), file_label, font=fs(20, bold=False), fill=(225, 222, 215), anchor="mm")
    img.alpha_composite(ov)
    sb_w = int((x1 - x0) * 0.205)
    sb = (x0, y0 + tb_h, x0 + sb_w, y1)
    grad_round(img, (sb[0], sb[1], sb[2] + 24, sb[3]), 0, PRIMARY_LT, PRIMARY_DK)
    grad_round(img, (sb[0], y1 - 24, sb[2], y1), 24, PRIMARY_DK, PRIMARY_DK)
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
    bx = sb[0] + 26
    od.text((bx, sb[1] + 30), "BABY COMMAND", font=fs(19), fill=GOLD_HI, anchor="lt")
    od.text((bx, sb[1] + 56), "21-tab system", font=fs(15, bold=False), fill=(170, 200, 192), anchor="lt")
    od.line((sb[0] + 20, sb[1] + 88, sb[2] - 16, sb[1] + 88), fill=(255, 255, 255, 40), width=1)
    list_top = sb[1] + 100
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
    ("BABY'S AGE", "125 days", "~18 weeks"),
    ("FEEDINGS TODAY", "7", "goal 8"),
    ("SLEEP TODAY", "14.0 h", "goal 15h"),
    ("DIAPERS TODAY", "6", "wet + dirty"),
    ("NEXT APPT", "3 days", "lactation f/u"),
    ("GROWTH LOGS", "5", "62nd %ile"),
    ("MONTHLY BUDGET", "$650", "set in Settings"),
    ("MILESTONES", "64%", "9 of 14"),
    ("VACCINES", "75%", "3 of 4 done"),
    ("BUDGET LEFT", "$50", "of $650"),
    ("GOALS", "48%", "family goals"),
    ("ROUTINE", "68%", "today"),
]


def content_dashboard(img, cbox):
    x0, y0, x1, y1 = cbox; pad = 30
    d = ImageDraw.Draw(img); d.rectangle(cbox, fill=BG)
    d.text((x0 + pad, y0 + 22), "Executive Baby Dashboard", font=fs(33), fill=PRIMARY, anchor="lt")
    d.text((x0 + pad, y0 + 64), "Baby Rose · 4 months  ·  your whole day, gently organized", font=fs(19, bold=False), fill=TEXT_MUTED, anchor="lt")
    d.rounded_rectangle((x1 - pad - 150, y0 + 26, x1 - pad, y0 + 62), radius=18, fill=MINT_BG)
    d.text((x1 - pad - 75, y0 + 44), "● live", font=fs(18), fill=PRIMARY, anchor="mm")
    gx = x0 + pad; gy = y0 + 98; gw = (x1 - x0 - 2 * pad); gap = 14
    kw = (gw - 5 * gap) / 6; kh = 116
    for i, (lab, val, sub) in enumerate(KPIS):
        r, ci = divmod(i, 6)
        kx = gx + ci * (kw + gap); ky = gy + r * (kh + gap)
        d.rounded_rectangle((kx, ky, kx + kw, ky + kh), radius=12, fill=WHITE, outline=GRID, width=2)
        d.rounded_rectangle((kx + 12, ky, kx + kw - 12, ky + 5), radius=2, fill=GOLD_LT)
        d.text((kx + 14, ky + 16), lab, font=fs(12), fill=ACCENT, anchor="lt")
        vf = fit_font(d, val, kw - 28, 30)
        d.text((kx + 14, ky + 58), val, font=vf, fill=PRIMARY, anchor="lm")
        d.text((kx + 14, ky + 96), sub, font=fs(12, bold=False), fill=TEXT_MUTED, anchor="lm")
    cy_top = gy + 2 * (kh + gap) + 18
    d.text((gx, cy_top), "DAILY RHYTHMS · GROWTH · BUDGET", font=fs(20), fill=ACCENT, anchor="lt")
    panels_y = cy_top + 34; panel_h = (y1 - panels_y - pad); pw = (gw - 3 * gap) / 4
    # feeding line
    px = gx
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Feedings / Day", font=fs(17), fill=ACCENT, anchor="lt")
    mini_lines(d, (px + 26, panels_y + 56, px + pw - 20, panels_y + panel_h - 40), [[8, 7, 8, 6, 7, 8, 7]], [PRIMARY])
    d.text((px + pw / 2, panels_y + panel_h - 22), "7-day avg 7.3", font=fs(15, bold=False), fill=TEXT_MUTED, anchor="mm")
    # sleep line
    px = gx + (pw + gap)
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Sleep Hours / Day", font=fs(17), fill=ACCENT, anchor="lt")
    mini_lines(d, (px + 26, panels_y + 56, px + pw - 20, panels_y + panel_h - 40), [[13.5, 14, 14, 13, 14, 13, 14]], [HIGHLIGHT])
    d.text((px + pw / 2, panels_y + panel_h - 22), "7-day avg 13.6h", font=fs(15, bold=False), fill=TEXT_MUTED, anchor="mm")
    # growth line
    px = gx + 2 * (pw + gap)
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Weight (lb)", font=fs(17), fill=ACCENT, anchor="lt")
    mini_lines(d, (px + 26, panels_y + 56, px + pw - 20, panels_y + panel_h - 40), [[7.4, 7.1, 9.2, 11.6, 14.2]], [ACCENT])
    d.text((px + pw / 2, panels_y + panel_h - 22), "birth → 4 mo", font=fs(15, bold=False), fill=TEXT_MUTED, anchor="mm")
    # budget donut
    px = gx + 3 * (pw + gap)
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Baby Spending", font=fs(17), fill=ACCENT, anchor="lt")
    donut(d, px + pw * 0.42, panels_y + panel_h * 0.52, min(panel_h * 0.28, pw * 0.27),
          [(25, PRIMARY), (15, ACCENT), (15, HIGHLIGHT), (12, SURFACE), (33, (170, 150, 120))], "$600", "spent")
    legend(d, px + pw * 0.04, panels_y + panel_h - 104, [(PRIMARY, "Formula 25%"), (ACCENT, "Diapers 15%"), (HIGHLIGHT, "Insurance 15%")], 15, 30)


def _basic_table(img, cbox, title, subtitle, headers, colf, rows, active_lm=True,
                 status_col=None, status_map=None, highlight_col=None):
    x0, y0, x1, y1 = cbox; pad = 30
    d = ImageDraw.Draw(img); d.rectangle(cbox, fill=BG)
    d.text((x0 + pad, y0 + 22), title, font=fs(32), fill=PRIMARY, anchor="lt")
    d.text((x0 + pad, y0 + 62), subtitle, font=fs(18, bold=False), fill=TEXT_MUTED, anchor="lt")
    tx0, tx1 = x0 + pad, x1 - pad
    ty = y0 + 104
    colx = [tx0 + (tx1 - tx0) * f for f in colf]
    hdr_h = 44
    grad_round(img, (tx0, ty, tx1, ty + hdr_h), 8, PRIMARY_LT, PRIMARY_DK)
    for i, h in enumerate(headers):
        anc = "lm" if i == 0 else "mm"
        d.text((colx[i] + (14 if i == 0 else 0), ty + hdr_h / 2), h, font=fs(15), fill=WHITE, anchor=anc)
    rh = (y1 - pad - (ty + hdr_h)) / len(rows)
    for i, row in enumerate(rows):
        ry = ty + hdr_h + i * rh
        hl = highlight_col is not None and str(row[highlight_col]).strip() not in ("", "—")
        if hl:
            d.rectangle((tx0, ry, tx1, ry + rh), fill=MINT_BG)
        elif i % 2:
            d.rectangle((tx0, ry, tx1, ry + rh), fill=ROW_ALT)
        for ci, val in enumerate(row):
            anc = "lm" if ci == 0 else "mm"
            hx = colx[ci] + (14 if ci == 0 else 0)
            if status_map is not None and ci == status_col:
                bg, fg = status_map.get(str(val), ((235, 230, 222), TEXT_MUTED))
                d.rounded_rectangle((hx - 52, ry + rh / 2 - 15, hx + 52, ry + rh / 2 + 15), radius=14, fill=bg)
                d.text((hx, ry + rh / 2), str(val), font=fs(15), fill=fg, anchor="mm")
            else:
                col = PRIMARY if ci == 0 else TEXT
                d.text((hx, ry + rh / 2), str(val), font=fs(17) if ci == 0 else fs(16, bold=False), fill=col, anchor=anc)


def content_feeding(img, cbox):
    rows = [
        ("6:00 AM", "Breast (L)", "15 min", "Today"),
        ("9:00 AM", "Breast (R)", "15 min", "Today"),
        ("12:00 PM", "Bottle - BM", "4 oz", "Today"),
        ("3:00 PM", "Bottle - Formula", "4 oz", "Today"),
        ("6:00 PM", "Breast (L)", "15 min", "Today"),
        ("8:30 PM", "Breast (R)", "15 min", "Today"),
        ("11:00 PM", "Bottle - BM", "4 oz", "Today"),
        ("6:15 AM", "Breast (L)", "15 min", "Yesterday"),
        ("9:20 AM", "Bottle - Formula", "4 oz", "Yesterday"),
        ("12:30 PM", "Breast (R)", "15 min", "Yesterday"),
    ]
    _basic_table(img, cbox, "Feeding Command Center",
                 "Every feed logged — 7 today (goal 8) · daily counts chart themselves",
                 ["TIME", "TYPE", "AMOUNT", "DAY"], [0.0, 0.30, 0.62, 0.82], rows)


def content_diaper(img, cbox):
    rows = [
        ("6:15 AM", "Wet", "Today"),
        ("9:15 AM", "Dirty", "Today"),
        ("12:15 PM", "Mixed", "Today"),
        ("3:15 PM", "Wet", "Today"),
        ("6:15 PM", "Dirty", "Today"),
        ("9:00 PM", "Mixed", "Today"),
        ("6:20 AM", "Wet", "Yesterday"),
        ("9:30 AM", "Dirty", "Yesterday"),
    ]
    _basic_table(img, cbox, "Diaper Tracker",
                 "Wet, dirty & mixed — 6 changes today · totals & trends update automatically",
                 ["TIME", "TYPE", "DAY"], [0.0, 0.42, 0.74], rows,
                 status_col=1, status_map={"Wet": (MINT_BG, PRIMARY), "Dirty": (WARN_BG, ACCENT), "Mixed": ((235, 230, 222), TEXT_MUTED)})


def content_growth(img, cbox):
    rows = [
        ("Birth", "7.4 lb", "20.0 in", "13.5 in", "—"),
        ("2 weeks", "9.2 lb", "21.5 in", "14.3 in", "55th %ile"),
        ("2 months", "11.6 lb", "22.8 in", "15.1 in", "60th %ile"),
        ("4 months", "14.2 lb", "24.4 in", "15.9 in", "62nd %ile"),
    ]
    _basic_table(img, cbox, "Growth Tracker",
                 "Weight, length & head circumference over time — with a growth chart",
                 ["VISIT", "WEIGHT", "LENGTH", "HEAD", "PERCENTILE"], [0.0, 0.30, 0.48, 0.64, 0.80], rows)


def content_milestones(img, cbox):
    rows = [
        ("First real smile", "Social", "Achieved", "✓"),
        ("Laughs out loud", "Social", "Achieved", "✓"),
        ("Holds head up", "Motor", "Achieved", "✓"),
        ("Rolls tummy to back", "Motor", "Achieved", "✓"),
        ("Grabs toys", "Motor", "Achieved", "✓"),
        ("6-hour sleep stretch", "Sleep", "Achieved", "✓"),
        ("Sits unassisted", "Motor", "—", ""),
        ("First tooth", "Teeth", "—", ""),
        ("First word", "Language", "—", ""),
        ("Pulls to stand", "Walking", "—", ""),
    ]
    _basic_table(img, cbox, "Milestone Tracker",
                 "Every first, captured — 9 of 14 logged · achieved rows glow mint",
                 ["MILESTONE", "CATEGORY", "STATUS", ""], [0.0, 0.44, 0.68, 0.90], rows,
                 highlight_col=3)


# ---------- renders ----------

def render_hero(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=640)
    d = ImageDraw.Draw(img)
    bottle_crest(img, SIZE // 2, 132, r=56)
    pill(img, SIZE // 2, 256, "THE ULTIMATE BABY CARE & FAMILY ORGANIZATION SYSTEM", font=fs(26), pad_x=42, pad_y=20)
    wordmark(img, SIZE // 2, 400, "BABY COMMAND CENTER", 116, max_w=1780)
    gold_divider(img, SIZE // 2, 500, width=520)
    tc(d, (SIZE // 2, 550), "Feeding, sleep, diapers, growth, health & memories — your whole day in one gentle place.",
       fs(25, bold=False), (224, 213, 190))
    chips = [("21", "GENTLE TABS"), ("AUTO", "DAILY TOTALS"), ("2-in-1", "EXCEL + SHEETS")]
    cw = 420
    total = len(chips) * cw + (len(chips) - 1) * 32
    startx = (SIZE - total) // 2 + cw // 2
    for i, (b, s) in enumerate(chips):
        stat_chip(img, startx + i * (cw + 32), 704, b, s, w=cw)
    app_window(img, (70, 800, SIZE - 70, 1900), 0, content_dashboard)
    pill(img, SIZE // 2, SIZE - 52, "21 SHEETS · INSTANT DOWNLOAD · EXCEL + GOOGLE SHEETS",
         font=fs(33), pad_x=50, pad_y=24, star=True, grad=(PRIMARY_LT, PRIMARY_DK))
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_inside(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=400)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 120, "EVERYTHING INSIDE", font=fs(38), pad_x=54, pad_y=22)
    tc(d, (SIZE // 2, 238), "21 Gentle, Connected Tabs", fserif(58), WHITE)
    gold_divider(img, SIZE // 2, 308, width=520)
    tc(d, (SIZE // 2, 352), "Not a baby tracker — a complete family baby operating system that runs itself",
       fs(24, bold=False), (226, 214, 190))
    cards = [
        ("Baby Dashboard", "12 KPIs + 4 live charts"), ("Baby Profile", "care team & medical"),
        ("Daily Log", "today at a glance"), ("Feeding Center", "auto daily counts"),
        ("Sleep Tracker", "hours & trends"), ("Diaper Tracker", "wet / dirty totals"),
        ("Growth Tracker", "weight & length chart"), ("Milestones", "every first, by category"),
        ("Medical Center", "visits & vaccines"), ("Appointments", "countdowns"),
        ("Baby Budget", "12 cost categories"), ("Shopping & Supplies", "reorder alerts"),
        ("Baby Inventory", "sizes & gear"), ("Childcare", "sitters & daycare"),
        ("Routine Planner", "gentle rhythms"), ("Development", "play with a purpose"),
        ("Memory Book", "photo placeholders"), ("Travel with Baby", "packing lists"),
        ("Family Goals", "progress bars"), ("Analytics", "organization score"),
        ("Settings", "make it yours"),
    ]
    cols = 4
    margin = 90
    gx, gy = 22, 22
    cw = (SIZE - 2 * margin - (cols - 1) * gx) // cols
    top = 440
    rows_n = 6
    ch = (SIZE - top - 60 - (rows_n - 1) * gy) // rows_n
    for i, (title, sub) in enumerate(cards):
        r, ccol = divmod(i, cols)
        x = margin + ccol * (cw + gx); y = top + r * (ch + gy)
        shadow(img, (x, y, x + cw, y + ch), 14, 13, 42, 9)
        ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
        od.rounded_rectangle((x, y, x + cw, y + ch), radius=14, fill=WHITE, outline=(232, 224, 208), width=2)
        od.rounded_rectangle((x, y, x + 7, y + ch), radius=3, fill=GOLD_LT)
        od.rectangle((x + 3, y, x + 7, y + ch), fill=GOLD_LT)
        cyc = y + ch // 2; bx = x + 42
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
    pill(img, SIZE // 2, 116, "EXECUTIVE BABY DASHBOARD", font=fs(34), pad_x=50, pad_y=22)
    tc(d, (SIZE // 2, 232), "Your Whole Day, Gently Organized", fserif(52), WHITE)
    tc(d, (SIZE // 2, 300), "12 live KPIs + feeding, sleep, growth & budget charts — all auto-updating",
       fs(24, bold=False), (226, 214, 190))
    app_window(img, (70, 400, SIZE - 70, SIZE - 70), 0, content_dashboard)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_trackers(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=300)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 110, "THE DAILY TRACKERS", font=fs(36), pad_x=52, pad_y=22)
    tc(d, (SIZE // 2, 224), "Feeding & Diapers, Effortlessly Logged", fserif(46), WHITE)
    app_window(img, (60, 330, SIZE - 60, 1150), 3, content_feeding)
    app_window(img, (60, 1180, SIZE - 60, SIZE - 60), 5, content_diaper)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_growth_miles(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=300)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 110, "GROWTH & MILESTONES", font=fs(36), pad_x=52, pad_y=22)
    tc(d, (SIZE // 2, 224), "Watch Them Grow — And Never Forget A First", fserif(42), WHITE)
    app_window(img, (60, 330, SIZE - 60, 1080), 6, content_growth)
    app_window(img, (60, 1110, SIZE - 60, SIZE - 60), 7, content_milestones)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_mobile(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=400)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 130, "WORKS EVERYWHERE", font=fs(38), pad_x=54, pad_y=22)
    tc(d, (SIZE // 2, 250), "Excel · Google Sheets · Mobile", fserif(56), WHITE)
    tc(d, (SIZE // 2, 320), "Log a feed one-handed at 3am — everyone caring for baby stays in sync",
       fs(24, bold=False), (226, 214, 190))
    px, py = SIZE // 2, 1300
    pw, ph = 640, 1230
    phone = (px - pw // 2, py - ph // 2, px + pw // 2, py + ph // 2)
    shadow(img, phone, 64, 50, 110, 24)
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
    od.rounded_rectangle(phone, radius=64, fill=(26, 26, 30))
    bez = 22
    screen = (phone[0] + bez, phone[1] + bez + 30, phone[2] - bez, phone[3] - bez - 30)
    od.rounded_rectangle(screen, radius=44, fill=BG)
    od.rounded_rectangle((px - 95, phone[1] + 16, px + 95, phone[1] + 50), radius=18, fill=(14, 14, 18))
    img.alpha_composite(ov)
    sx0, sy0, sx1, sy1 = screen
    grad_round(img, (sx0, sy0, sx1, sy0 + 110), 44, PRIMARY_LT, PRIMARY_DK)
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
    od.rectangle((sx0, sy0 + 106, sx1, sy0 + 110), fill=GOLD_LT)
    od.text(((sx0 + sx1) // 2, sy0 + 56), "Baby Command", font=fserif(36), fill=GOLD_HI, anchor="mm")
    y = sy0 + 150
    cards = [("FEEDINGS TODAY", "7", PRIMARY), ("SLEEP TODAY", "14.0 h", PRIMARY),
             ("DIAPERS TODAY", "6", ACCENT), ("NEXT APPT", "3 days", PRIMARY)]
    for lab, val, col in cards:
        cb = (sx0 + 30, y, sx1 - 30, y + 135)
        od.rounded_rectangle(cb, radius=18, fill=WHITE, outline=GOLD_LT, width=2)
        od.rounded_rectangle((cb[0] + 20, y, cb[2] - 20, y + 5), radius=2, fill=GOLD_LT)
        od.text((cb[0] + 26, y + 34), lab, font=fs(22), fill=ACCENT, anchor="lt")
        vf = fit_font(od, val, sx1 - sx0 - 110, 46)
        od.text((cb[0] + 26, y + 94), val, font=vf, fill=col, anchor="lm")
        y += 152
    od.text((sx0 + 40, y + 16), "TODAY'S CARE", font=fs(22), fill=ACCENT, anchor="lt")
    y += 52
    for lab, state in [("Morning feed & cuddle", True), ("Tummy time (3x)", True),
                       ("Bath night", False), ("Restock diaper bag", False)]:
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
        ("04_trackers.png", render_trackers),
        ("05_growth_miles.png", render_growth_miles),
        ("06_mobile.png", render_mobile),
    ]
    for name, fn in targets:
        fn(os.path.join(out_dir, name))
        print(f"  ✓ {name}")
    print(f"Wrote {len(targets)} images to {out_dir}")


if __name__ == "__main__":
    main()

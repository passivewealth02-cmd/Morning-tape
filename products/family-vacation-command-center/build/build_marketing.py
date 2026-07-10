"""Marketing image set for Family Vacation Command Center™ (6 images, 2000x2000).

Dense app-screenshot marketing mirroring the real workbook: a left sidebar of
all 24 tabs, the REAL computed KPI numbers, and fully populated tables/charts.

  01_hero.png       - branded hero + live family travel dashboard
  02_inside.png     - "everything inside — 24 powerful tabs"
  03_budget.png     - budget command center (planned vs actual, cost per person)
  04_packing.png    - family packing command center
  05_plan.png       - itinerary + reservation organizer
  06_mobile.png     - mobile preview

Run: python3 build_marketing.py
"""
from __future__ import annotations
import datetime as dt
import math
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
GREEN = (32, 120, 96)
SKY = (94, 160, 200)
SKY_DK = (58, 118, 158)
SIZE = 2000

SANS_B = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
SANS_R = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
SERIF_B = "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf"

TABS = ["Dashboard", "Family Profile", "Master Trip", "Budget", "Savings", "Itinerary",
        "Transport", "Hotels", "Packing", "Kids", "Meals", "Activities", "Reservations",
        "Documents", "Emergency", "Road Trip", "Responsibilities", "Points", "Photo Vault",
        "Journal", "Post-Trip", "Wishlist", "Analytics", "Settings"]

FILE_LABEL = "Family_Vacation_Command_Center.xlsx — The Rivera Family · Walt Disney World"


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


def globe_crest(c, cx, cy, r=56, glow=True):
    """Brand crest with a globe + travel path."""
    if glow:
        radial_glow(c, cx, cy, int(r * 2.1), GOLD_HI, 90)
    grad_round(c, (cx - r, cy - r, cx + r, cy + r), 22, PRIMARY_LT, PRIMARY_DK, outline=GOLD_LT, width=4)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0)); d = ImageDraw.Draw(ov)
    d.rounded_rectangle((cx - r + 10, cy - r + 10, cx + r - 10, cy + r - 10), radius=16, outline=GOLD_HI, width=2)
    c.alpha_composite(ov)
    # globe
    gr = r * 0.6
    grad_round(c, (cx - gr, cy - gr, cx + gr, cy + gr), int(gr), SKY, SKY_DK, outline=GOLD_HI, width=3)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0)); d = ImageDraw.Draw(ov)
    # meridians / parallels
    d.ellipse((cx - gr * 0.42, cy - gr, cx + gr * 0.42, cy + gr), outline=(255, 255, 255, 120), width=2)
    d.line((cx - gr, cy, cx + gr, cy), fill=(255, 255, 255, 120), width=2)
    d.arc((cx - gr, cy - gr * 0.5, cx + gr, cy + gr * 1.5), 180, 360, fill=(255, 255, 255, 90), width=2)
    d.arc((cx - gr, cy - gr * 1.5, cx + gr, cy + gr * 0.5), 0, 180, fill=(255, 255, 255, 90), width=2)
    # little land blobs
    d.ellipse((cx - gr * 0.5, cy - gr * 0.4, cx - gr * 0.05, cy + gr * 0.05), fill=(120, 200, 150, 150))
    d.ellipse((cx + gr * 0.1, cy + gr * 0.05, cx + gr * 0.55, cy + gr * 0.5), fill=(120, 200, 150, 150))
    c.alpha_composite(ov)
    # dashed travel arc + plane dot over the crest
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0)); d = ImageDraw.Draw(ov)
    d.arc((cx - r * 0.7, cy - r * 0.85, cx + r * 0.7, cy + r * 0.2), 200, 340, fill=GOLD_HI, width=3)
    d.ellipse((cx + r * 0.5, cy - r * 0.5, cx + r * 0.62, cy - r * 0.38), fill=WHITE)
    c.alpha_composite(ov)


def stat_chip(c, cx, cy, big, small, w=400, h=150):
    box = (cx - w // 2, cy - h // 2, cx + w // 2, cy + h // 2)
    shadow(c, box, 20, 24, 80, 16)
    grad_round(c, box, 20, PRIMARY_LT, PRIMARY_DK, outline=GOLD_LT, width=3)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0)); d = ImageDraw.Draw(ov)
    d.rounded_rectangle((box[0] + 18, box[1] + 12, box[2] - 18, box[1] + 18), radius=3, fill=GOLD_HI)
    d.text((cx, cy - h * 0.16), big, font=fserif(48), fill=GOLD_HI, anchor="mm")
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


def hbars(img, d, box, items, color=(HIGHLIGHT, (70, 200, 165)), labelcol=ACCENT):
    x0, y0, x1, y1 = box; n = len(items); rowh = min((y1 - y0) / n, 110); bw_max = (x1 - x0) - 200
    y0 = y0 + max(((y1 - y0) - rowh * n) / 2, 0)
    for i, (lab, frac, vlabel) in enumerate(items):
        yy = y0 + rowh * i + rowh * 0.16
        d.text((x0, yy), lab, font=fs(18, bold=False), fill=TEXT, anchor="lt")
        d.rounded_rectangle((x0, yy + 30, x0 + bw_max, yy + 56), radius=13, fill=(236, 230, 220))
        grad_round(img, (x0, yy + 30, x0 + bw_max * frac, yy + 56), 13, color[0], color[1])
        d.text((x0 + bw_max + 14, yy + 43), vlabel, font=fs(18), fill=labelcol, anchor="lm")


def fit_font(d, text, max_w, start, serif=True):
    s = start; f = fserif(s) if serif else fs(s)
    while s > 12 and d.textlength(text, font=f) > max_w:
        s -= 1; f = fserif(s) if serif else fs(s)
    return f


def app_window(img, box, active_idx, content_fn, file_label=FILE_LABEL):
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
    od.text(((x0 + x1) / 2, y0 + tb_h // 2), file_label, font=fs(19, bold=False), fill=(225, 222, 215), anchor="mm")
    img.alpha_composite(ov)
    sb_w = int((x1 - x0) * 0.205)
    sb = (x0, y0 + tb_h, x0 + sb_w, y1)
    grad_round(img, (sb[0], sb[1], sb[2] + 24, sb[3]), 0, PRIMARY_LT, PRIMARY_DK)
    grad_round(img, (sb[0], y1 - 24, sb[2], y1), 24, PRIMARY_DK, PRIMARY_DK)
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
    bx = sb[0] + 26
    od.text((bx, sb[1] + 26), "FAMILY TRIP", font=fs(18), fill=GOLD_HI, anchor="lt")
    od.text((bx, sb[1] + 50), "24-tab system", font=fs(14, bold=False), fill=(170, 200, 192), anchor="lt")
    od.line((sb[0] + 20, sb[1] + 78, sb[2] - 16, sb[1] + 78), fill=(255, 255, 255, 40), width=1)
    list_top = sb[1] + 86
    rowh = (y1 - 20 - list_top) / len(TABS)
    palette = [HIGHLIGHT, GOLD_HI, SURFACE, (150, 200, 190)]
    for i, name in enumerate(TABS):
        ry = list_top + i * rowh
        if i == active_idx:
            od.rounded_rectangle((sb[0] + 12, ry + 1, sb[2] - 10, ry + rowh - 1), radius=6, fill=(255, 255, 255, 235))
            od.rounded_rectangle((sb[0] + 12, ry + 1, sb[0] + 18, ry + rowh - 1), radius=3, fill=GOLD_HI)
            dotc = PRIMARY; txtc = PRIMARY; font = fs(15)
        else:
            dotc = palette[i % len(palette)]; txtc = (214, 226, 222); font = fs(14, bold=False)
        cyr = ry + rowh / 2
        od.ellipse((sb[0] + 26, cyr - 4, sb[0] + 34, cyr + 4), fill=dotc)
        od.text((sb[0] + 46, cyr), name, font=font, fill=txtc, anchor="lm")
    img.alpha_composite(ov)
    cbox = (sb[2] + 1, y0 + tb_h + 4, x1, y1)
    content_fn(img, cbox)


KPIS = [
    ("TRIP COUNTDOWN", "45", "days to go"),
    ("TOTAL BUDGET", "$8,500", "5 travelers"),
    ("MONEY SAVED", "$6,800", "of $8,500"),
    ("BUDGET REMAINING", "$3,300", "after bookings"),
    ("TRAVELERS", "5", "family"),
    ("DAYS TRAVELING", "7", "nights"),
    ("PACKING COMPLETE", "72%", "23 of 32"),
    ("BOOKINGS DONE", "92%", "11 of 12"),
    ("DOCUMENTS READY", "83%", "5 of 6"),
    ("ACTIVITIES PLANNED", "14", "on the list"),
    ("COST / PERSON", "$1,700", "all-in"),
    ("READINESS SCORE", "82%", "trip ready"),
]


def content_dashboard(img, cbox):
    x0, y0, x1, y1 = cbox; pad = 30
    d = ImageDraw.Draw(img); d.rectangle(cbox, fill=BG)
    d.text((x0 + pad, y0 + 22), "Family Travel Dashboard", font=fs(32), fill=PRIMARY, anchor="lt")
    d.text((x0 + pad, y0 + 64), "The Rivera Family · Walt Disney World  ·  your whole trip, automatically organized", font=fs(19, bold=False), fill=TEXT_MUTED, anchor="lt")
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
    d.text((gx, cy_top), "BUDGET · SAVINGS · READINESS · SPENDING", font=fs(20), fill=ACCENT, anchor="lt")
    panels_y = cy_top + 34; panel_h = (y1 - panels_y - pad); pw = (gw - 3 * gap) / 4
    # budget donut
    px = gx
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Budget by Category", font=fs(17), fill=ACCENT, anchor="lt")
    donut(d, px + pw * 0.42, panels_y + panel_h * 0.52, min(panel_h * 0.28, pw * 0.27),
          [(28, PRIMARY), (26, ACCENT), (15, HIGHLIGHT), (14, SURFACE), (17, (170, 150, 120))], "$8.5K", "planned")
    legend(d, px + pw * 0.04, panels_y + panel_h - 104, [(PRIMARY, "Hotels 28%"), (ACCENT, "Flights 26%"), (HIGHLIGHT, "Food 15%")], 15, 30)
    # savings growth bars
    px = gx + (pw + gap)
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Vacation Fund ($)", font=fs(17), fill=ACCENT, anchor="lt")
    months = ["Feb", "Mar", "Apr", "May", "Jun", "Jul"]; vals = [1000, 2200, 3300, 4500, 5800, 6800]
    hbars(img, d, (px + 20, panels_y + 50, px + pw - 16, panels_y + panel_h - 16),
          [(m, v / 6800, f"{v/1000:.1f}K") for m, v in zip(months, vals)], color=(GOLD_HI, GOLD))
    # readiness bars
    px = gx + 2 * (pw + gap)
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Trip Readiness", font=fs(17), fill=ACCENT, anchor="lt")
    hbars(img, d, (px + 20, panels_y + 56, px + pw - 16, panels_y + panel_h - 20),
          [("Savings", 0.80, "80%"), ("Bookings", 0.92, "92%"), ("Packing", 0.72, "72%"), ("Documents", 0.83, "83%")])
    # planned vs actual donut
    px = gx + 3 * (pw + gap)
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Spent vs Remaining", font=fs(17), fill=ACCENT, anchor="lt")
    donut(d, px + pw * 0.42, panels_y + panel_h * 0.52, min(panel_h * 0.28, pw * 0.27),
          [(61, PRIMARY), (39, SURFACE)], "61%", "booked")
    legend(d, px + pw * 0.04, panels_y + panel_h - 76, [(PRIMARY, "Spent $5,200"), (SURFACE, "Left $3,300")], 15, 30)


def _table(img, cbox, title, subtitle, headers, colf, rows, total_row=None,
           status_col=None, status_map=None, hdr_top=104):
    x0, y0, x1, y1 = cbox; pad = 30
    d = ImageDraw.Draw(img); d.rectangle(cbox, fill=BG)
    d.text((x0 + pad, y0 + 22), title, font=fs(32), fill=PRIMARY, anchor="lt")
    d.text((x0 + pad, y0 + 62), subtitle, font=fs(18, bold=False), fill=TEXT_MUTED, anchor="lt")
    tx0, tx1 = x0 + pad, x1 - pad
    ty = y0 + hdr_top
    colx = [tx0 + (tx1 - tx0) * f for f in colf]
    hdr_h = 42
    grad_round(img, (tx0, ty, tx1, ty + hdr_h), 8, PRIMARY_LT, PRIMARY_DK)
    for i, h in enumerate(headers):
        anc = "lm" if i == 0 else "mm"
        d.text((colx[i] + (14 if i == 0 else 0), ty + hdr_h / 2), h, font=fs(15), fill=WHITE, anchor=anc)
    nrows = len(rows) + (1 if total_row else 0)
    rh = (y1 - pad - (ty + hdr_h)) / nrows
    for i, row in enumerate(rows):
        ry = ty + hdr_h + i * rh
        if i % 2:
            d.rectangle((tx0, ry, tx1, ry + rh), fill=ROW_ALT)
        for ci, val in enumerate(row):
            anc = "lm" if ci == 0 else "mm"
            hx = colx[ci] + (14 if ci == 0 else 0)
            sval = str(val)
            if status_map is not None and ci == status_col:
                bg, fg = status_map.get(sval, ((235, 230, 222), TEXT_MUTED))
                d.rounded_rectangle((hx - 62, ry + rh / 2 - 15, hx + 62, ry + rh / 2 + 15), radius=14, fill=bg)
                d.text((hx, ry + rh / 2), sval, font=fs(15), fill=fg, anchor="mm")
            else:
                col = PRIMARY if ci == 0 else TEXT
                d.text((hx, ry + rh / 2), sval, font=fs(16) if ci == 0 else fs(15, bold=False), fill=col, anchor=anc)
    if total_row:
        ry = ty + hdr_h + len(rows) * rh
        d.rectangle((tx0, ry, tx1, ry + rh), fill=SURFACE)
        for ci, val in enumerate(total_row):
            anc = "lm" if ci == 0 else "mm"
            hx = colx[ci] + (14 if ci == 0 else 0)
            if val != "":
                d.text((hx, ry + rh / 2), str(val), font=fs(16), fill=PRIMARY, anchor=anc)


def content_budget(img, cbox):
    rows = [
        ("Flights", "$2,200", "$2,200", "$0"),
        ("Hotels", "$2,400", "$2,400", "$0"),
        ("Food & Dining", "$1,300", "$0", "$1,300"),
        ("Activities", "$1,200", "$600", "$600"),
        ("Rental Car", "$450", "$0", "$450"),
        ("Shopping", "$300", "$0", "$300"),
        ("Insurance", "$250", "$0", "$250"),
        ("Emergency Fund", "$200", "$0", "$200"),
        ("Transportation", "$200", "$0", "$200"),
    ]
    _table(img, cbox, "Family Vacation Budget Command Center",
           "Every dollar planned — budget vs actual, cost per person & per day, live.  $1,700/person · $1,214/day",
           ["CATEGORY", "PLANNED", "ACTUAL", "REMAINING"],
           [0.0, 0.42, 0.64, 0.88], rows,
           total_row=("TOTAL BUDGET", "$8,500", "$5,200", "$3,300"))


def content_packing(img, cbox):
    rows = [
        ("Passports & IDs", "Documents", "Marcus", "5", "Yes"),
        ("Park tickets (mobile)", "Documents", "Elena", "5", "Yes"),
        ("Family medicine kit", "Medicine", "Elena", "1", "Yes"),
        ("Mateo's allergy meds", "Medicine", "Elena", "1", "Yes"),
        ("Sunscreen SPF 50", "Toiletries", "Elena", "3", "Yes"),
        ("Stroller (Mateo)", "Kids", "Marcus", "1", "Yes"),
        ("Kids tablets + cases", "Kids", "Sofia", "2", "Yes"),
        ("Swimsuits", "Clothing", "Each", "10", "Yes"),
        ("Camera + SD cards", "Electronics", "Sofia", "1", "No"),
        ("Rain ponchos", "Clothing", "Marcus", "5", "No"),
        ("Neck pillows", "Comfort", "Each", "5", "No"),
        ("Snacks for travel", "Kids", "Elena", "1", "No"),
    ]
    _table(img, cbox, "Family Packing Command Center",
           "Nothing forgotten — assign every item and check it off.  72% packed (23 of 32)",
           ["ITEM", "CATEGORY", "ASSIGNED", "QTY", "PACKED?"],
           [0.0, 0.40, 0.60, 0.76, 0.90], rows,
           status_col=4, status_map={"Yes": (MINT_BG, PRIMARY), "No": (WARN_BG, ACCENT)})


def content_itinerary(img, cbox):
    rows = [
        ("Day 1", "Travel & Check-In", "Disney Springs", "Quick-service", "Arrival — easy"),
        ("Day 2", "Magic Kingdom", "Parade & lunch", "Character dining", "Fireworks night"),
        ("Day 3", "EPCOT", "World Showcase", "Space 220 lunch", "Nap for Mateo"),
        ("Day 4", "Rest & Pool Day", "Mini-golf", "Cook in room", "Recharge"),
        ("Day 5", "Hollywood Studios", "Toy Story Land", "Sci-Fi Diner", "Fantasmic!"),
        ("Day 6", "Animal Kingdom", "Safari midday", "Tusker breakfast", "Water bottles!"),
        ("Day 7", "Departure", "Souvenirs & pack", "Last breakfast", "Bags down 9am"),
    ]
    _table(img, cbox, "Itinerary Command Center",
           "Day by day — every morning, afternoon & evening planned, with meals & notes",
           ["DAY", "THEME", "AFTERNOON", "MEALS / RES.", "NOTES"],
           [0.0, 0.20, 0.44, 0.64, 0.84], rows)


def content_reservations(img, cbox):
    rows = [
        ("Flight", "DL-8KX2P9", "Delta", "Paid", "Yes"),
        ("Hotel", "WDW-CB-55219", "Disney", "Deposit", "Yes"),
        ("Park Ticket", "WDW-TIX-7742", "Disney", "Paid", "Yes"),
        ("Restaurant", "CRT-6PM", "Disney", "Paid", "Yes"),
        ("Restaurant", "S220-1230", "Disney", "Paid", "Yes"),
        ("Activity", "BBB-0946", "Disney", "Paid", "Yes"),
        ("Rental Car", "ENT-77341", "Enterprise", "Pending", "Yes"),
        ("Parking", "—", "Resort", "Pending", "No"),
    ]
    _table(img, cbox, "Reservation Organizer",
           "Nothing double-booked — confirmations, payment status & cancel deadlines",
           ["TYPE", "CONFIRMATION", "WHERE", "PAYMENT", "OK?"],
           [0.0, 0.28, 0.56, 0.72, 0.90], rows,
           status_col=3, status_map={"Paid": (MINT_BG, PRIMARY), "Deposit": (WARN_BG, ACCENT),
                                     "Pending": (RED_BG, DANGER)})


# ---------- renders ----------

def render_hero(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=640)
    d = ImageDraw.Draw(img)
    globe_crest(img, SIZE // 2, 132, r=56)
    pill(img, SIZE // 2, 256, "THE FAMILY TRAVEL OPERATING SYSTEM", font=fs(25), pad_x=42, pad_y=20)
    wordmark(img, SIZE // 2, 400, "FAMILY VACATION COMMAND CENTER", 88, max_w=1860)
    gold_divider(img, SIZE // 2, 500, width=520)
    tc(d, (SIZE // 2, 550), "Plan, budget, pack, organize & remember every trip — from the first idea to the final photo.",
       fs(22, bold=False), (224, 213, 190))
    chips = [("24", "POWERFUL TABS"), ("AUTO", "BUDGET + READINESS"), ("2-in-1", "EXCEL + SHEETS")]
    cw = 420
    total = len(chips) * cw + (len(chips) - 1) * 32
    startx = (SIZE - total) // 2 + cw // 2
    for i, (b, s) in enumerate(chips):
        stat_chip(img, startx + i * (cw + 32), 704, b, s, w=cw)
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
    tc(d, (SIZE // 2, 352), "Not a vacation checklist — a complete family travel operating system",
       fs(24, bold=False), (226, 214, 190))
    cards = [
        ("Travel Dashboard", "12 KPIs + live charts"), ("Family Profile", "sizes, allergies & more"),
        ("Master Trip Planner", "the big picture"), ("Budget Command", "plan vs actual"),
        ("Savings Planner", "grow the fund"), ("Itinerary", "day-by-day plan"),
        ("Flights & Transport", "every leg covered"), ("Hotels", "confirmations & dates"),
        ("Packing Command", "nothing forgotten"), ("Kids Organizer", "happy little travelers"),
        ("Meal Planner", "diets & reservations"), ("Activity Planner", "the fun list"),
        ("Reservations", "never double-booked"), ("Document Center", "papers in order"),
        ("Emergency Center", "contacts & health"), ("Road Trip Center", "route, fuel & stops"),
        ("Responsibility Board", "everyone helps"), ("Points & Rewards", "travel for less"),
        ("Photo & Memory Vault", "keep the magic"), ("Travel Journal", "one line a day"),
        ("Post-Trip Review", "learn & improve"), ("Future Wishlist", "the next adventure"),
        ("Travel Analytics", "readiness score"), ("Settings", "your lists & trip"),
    ]
    cols = 4
    margin = 88
    gx, gy = 22, 20
    cw = (SIZE - 2 * margin - (cols - 1) * gx) // cols
    top = 438
    rows_n = 6
    ch = (SIZE - top - 56 - (rows_n - 1) * gy) // rows_n
    for i, (title, sub) in enumerate(cards):
        r, ccol = divmod(i, cols)
        x = margin + ccol * (cw + gx); y = top + r * (ch + gy)
        shadow(img, (x, y, x + cw, y + ch), 14, 12, 40, 8)
        ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
        od.rounded_rectangle((x, y, x + cw, y + ch), radius=14, fill=WHITE, outline=(232, 224, 208), width=2)
        od.rectangle((x + 3, y + 8, x + 7, y + ch - 8), fill=GOLD_LT)
        cyc = y + ch // 2; bx = x + 44
        od.ellipse((bx - 22, cyc - 22, bx + 22, cyc + 22), fill=PRIMARY)
        od.text((bx, cyc), str(i + 1), font=fs(20), fill=GOLD_HI, anchor="mm")
        od.text((x + 80, cyc - 16), title, font=fs(20), fill=PRIMARY, anchor="lm")
        od.text((x + 80, cyc + 19), sub, font=fs(14, bold=False), fill=TEXT_MUTED, anchor="lm")
        img.alpha_composite(ov)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_budget(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=360)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 116, "BUDGET COMMAND CENTER", font=fs(32), pad_x=48, pad_y=22)
    tc(d, (SIZE // 2, 232), "Every Dollar, Planned & Tracked", fserif(50), WHITE)
    tc(d, (SIZE // 2, 300), "Budget vs actual, cost per person & per day — no more vacation money surprises",
       fs(23, bold=False), (226, 214, 190))
    app_window(img, (70, 400, SIZE - 70, SIZE - 70), 3, content_budget)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_packing(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=360)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 116, "PACKING COMMAND CENTER", font=fs(32), pad_x=48, pad_y=22)
    tc(d, (SIZE // 2, 232), "Nothing Left Behind", fserif(52), WHITE)
    tc(d, (SIZE // 2, 300), "Assign every item, track quantity & check it off — packing completion updates itself",
       fs(23, bold=False), (226, 214, 190))
    app_window(img, (70, 400, SIZE - 70, SIZE - 70), 8, content_packing)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_plan(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=300)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 110, "PLAN EVERY DAY", font=fs(36), pad_x=52, pad_y=22)
    tc(d, (SIZE // 2, 224), "The Whole Trip, Organized", fserif(46), WHITE)
    app_window(img, (60, 330, SIZE - 60, 1150), 5, content_itinerary)
    app_window(img, (60, 1180, SIZE - 60, SIZE - 60), 12, content_reservations)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_mobile(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=400)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 130, "WORKS EVERYWHERE", font=fs(38), pad_x=54, pad_y=22)
    tc(d, (SIZE // 2, 250), "Excel · Google Sheets · Mobile", fserif(56), WHITE)
    tc(d, (SIZE // 2, 320), "Check the budget from the airport — your whole trip in your pocket",
       fs(23, bold=False), (226, 214, 190))
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
    od.text(((sx0 + sx1) // 2, sy0 + 56), "Family Trip", font=fserif(32), fill=GOLD_HI, anchor="mm")
    y = sy0 + 150
    cards = [("TRIP COUNTDOWN", "45 days", PRIMARY), ("MONEY SAVED", "$6,800", ACCENT),
             ("PACKING COMPLETE", "72%", PRIMARY), ("READINESS SCORE", "82%", PRIMARY)]
    for lab, val, col in cards:
        cb = (sx0 + 30, y, sx1 - 30, y + 135)
        od.rounded_rectangle(cb, radius=18, fill=WHITE, outline=GOLD_LT, width=2)
        od.rounded_rectangle((cb[0] + 20, y, cb[2] - 20, y + 5), radius=2, fill=GOLD_LT)
        od.text((cb[0] + 26, y + 34), lab, font=fs(22), fill=ACCENT, anchor="lt")
        vf = fit_font(od, val, sx1 - sx0 - 110, 46)
        od.text((cb[0] + 26, y + 94), val, font=vf, fill=col, anchor="lm")
        y += 152
    od.text((sx0 + 40, y + 16), "BEFORE WE GO", font=fs(22), fill=ACCENT, anchor="lt")
    y += 52
    for lab, state in [("Confirm rental car", False), ("Finish travel insurance", False),
                       ("Pack camera + charger", True), ("Save July $1,000", True)]:
        col = HIGHLIGHT if state else BG
        od.ellipse((sx0 + 40, y + 6, sx0 + 78, y + 44), fill=col, outline=PRIMARY, width=3)
        if state:
            od.text((sx0 + 59, y + 24), "✓", font=fs(24), fill=PRIMARY, anchor="mm")
        od.text((sx0 + 96, y + 24), lab, font=fs(21, bold=False), fill=TEXT, anchor="lm")
        y += 64
    img.alpha_composite(ov)
    img.convert("RGB").save(out, "PNG", optimize=True)


def main():
    out_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "marketing")
    os.makedirs(out_dir, exist_ok=True)
    targets = [
        ("01_hero.png", render_hero),
        ("02_inside.png", render_inside),
        ("03_budget.png", render_budget),
        ("04_packing.png", render_packing),
        ("05_plan.png", render_plan),
        ("06_mobile.png", render_mobile),
    ]
    for name, fn in targets:
        fn(os.path.join(out_dir, name))
        print(f"  ✓ {name}")
    print(f"Wrote {len(targets)} images to {out_dir}")


if __name__ == "__main__":
    main()

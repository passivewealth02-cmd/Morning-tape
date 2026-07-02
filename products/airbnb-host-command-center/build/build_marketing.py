"""Marketing image set for Airbnb Host Command Center™ (6 images, 2000x2000).

Dense app-screenshot marketing mirroring the real workbook: a left sidebar of
all 19 tabs, the REAL computed KPI numbers, and fully populated tables/charts.

  01_hero.png       - branded hero + live host dashboard
  02_inside.png     - "everything inside — 19 powerful tabs"
  03_dashboard.png  - full executive host dashboard
  04_financial.png  - financial command center (live P&L)
  05_manage.png     - booking calendar + reservation manager
  06_mobile.png     - mobile preview

Run: python3 build_marketing.py
"""
from __future__ import annotations
import datetime as dt
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

TABS = ["Dashboard", "Property Profile", "Calendar", "Reservations", "Financial",
        "Pricing", "Guests", "Cleaning", "Maintenance", "Inventory", "Reviews",
        "Messages", "Suppliers", "Taxes", "Multi-Property", "Goals",
        "Improvements", "Analytics", "Settings"]

FILE_LABEL = "Airbnb_Host_Command_Center.xlsx — Summit Stays Co. · 3 properties"


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


def house_crest(c, cx, cy, r=60, glow=True):
    if glow:
        radial_glow(c, cx, cy, int(r * 2.1), GOLD_HI, 90)
    grad_round(c, (cx - r, cy - r, cx + r, cy + r), 22, PRIMARY_LT, PRIMARY_DK, outline=GOLD_LT, width=4)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0)); d = ImageDraw.Draw(ov)
    d.rounded_rectangle((cx - r + 10, cy - r + 10, cx + r - 10, cy + r - 10), radius=16, outline=GOLD_HI, width=2)
    # house: roof + body + door + windows
    roof_peak = (cx, cy - r * 0.52)
    roof_l = (cx - r * 0.52, cy - r * 0.06)
    roof_r = (cx + r * 0.52, cy - r * 0.06)
    d.polygon([roof_peak, roof_l, roof_r], fill=GOLD_HI)
    body = (cx - r * 0.38, cy - r * 0.06, cx + r * 0.38, cy + r * 0.5)
    d.rectangle(body, fill=GOLD)
    # door
    d.rectangle((cx - r * 0.1, cy + r * 0.12, cx + r * 0.1, cy + r * 0.5), fill=PRIMARY_DK)
    # windows
    for wx in (-r * 0.24, r * 0.24):
        d.rectangle((cx + wx - r * 0.09, cy + r * 0.02, cx + wx + r * 0.09, cy + r * 0.2), fill=PRIMARY_DK)
    # growth chevron (business signal) above roof
    d.line([(cx - r * 0.22, cy - r * 0.66), (cx, cy - r * 0.78), (cx + r * 0.22, cy - r * 0.62)], fill=HIGHLIGHT, width=4)
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


def hbars(img, d, box, items):
    x0, y0, x1, y1 = box; n = len(items); rowh = min((y1 - y0) / n, 110); bw_max = (x1 - x0) - 260
    y0 = y0 + max(((y1 - y0) - rowh * n) / 2, 0)
    for i, (lab, pct) in enumerate(items):
        yy = y0 + rowh * i + rowh * 0.16
        d.text((x0, yy), lab, font=fs(18, bold=False), fill=TEXT, anchor="lt")
        d.rounded_rectangle((x0, yy + 30, x0 + bw_max, yy + 56), radius=13, fill=(236, 230, 220))
        grad_round(img, (x0, yy + 30, x0 + bw_max * pct, yy + 56), 13, HIGHLIGHT, (70, 200, 165))
        d.text((x0 + bw_max + 14, yy + 43), f"{int(pct*100)}%", font=fs(18), fill=ACCENT, anchor="lm")


def hbars_money(img, d, box, items):
    x0, y0, x1, y1 = box; n = len(items); rowh = min((y1 - y0) / n, 110); bw_max = (x1 - x0) - 190
    y0 = y0 + max(((y1 - y0) - rowh * n) / 2, 0)
    for i, (lab, frac, vlabel) in enumerate(items):
        yy = y0 + rowh * i + rowh * 0.16
        d.text((x0, yy), lab, font=fs(18, bold=False), fill=TEXT, anchor="lt")
        d.rounded_rectangle((x0, yy + 30, x0 + bw_max, yy + 56), radius=13, fill=(236, 230, 220))
        grad_round(img, (x0, yy + 30, x0 + bw_max * frac, yy + 56), 13, PRIMARY_LT, PRIMARY_DK)
        d.text((x0 + bw_max + 14, yy + 43), vlabel, font=fs(18), fill=PRIMARY, anchor="lm")


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
    od.text(((x0 + x1) / 2, y0 + tb_h // 2), file_label, font=fs(20, bold=False), fill=(225, 222, 215), anchor="mm")
    img.alpha_composite(ov)
    sb_w = int((x1 - x0) * 0.205)
    sb = (x0, y0 + tb_h, x0 + sb_w, y1)
    grad_round(img, (sb[0], sb[1], sb[2] + 24, sb[3]), 0, PRIMARY_LT, PRIMARY_DK)
    grad_round(img, (sb[0], y1 - 24, sb[2], y1), 24, PRIMARY_DK, PRIMARY_DK)
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
    bx = sb[0] + 26
    od.text((bx, sb[1] + 30), "HOST OS", font=fs(19), fill=GOLD_HI, anchor="lt")
    od.text((bx, sb[1] + 56), "19-tab system", font=fs(15, bold=False), fill=(170, 200, 192), anchor="lt")
    od.line((sb[0] + 20, sb[1] + 88, sb[2] - 16, sb[1] + 88), fill=(255, 255, 255, 40), width=1)
    list_top = sb[1] + 100
    rowh = (y1 - 24 - list_top) / len(TABS)
    palette = [HIGHLIGHT, GOLD_HI, SURFACE, (150, 200, 190)]
    for i, name in enumerate(TABS):
        ry = list_top + i * rowh
        if i == active_idx:
            od.rounded_rectangle((sb[0] + 12, ry + 2, sb[2] - 10, ry + rowh - 2), radius=8, fill=(255, 255, 255, 235))
            od.rounded_rectangle((sb[0] + 12, ry + 2, sb[0] + 19, ry + rowh - 2), radius=3, fill=GOLD_HI)
            dotc = PRIMARY; txtc = PRIMARY; font = fs(17)
        else:
            dotc = palette[i % len(palette)]; txtc = (214, 226, 222); font = fs(16, bold=False)
        cyr = ry + rowh / 2
        od.ellipse((sb[0] + 30, cyr - 6, sb[0] + 42, cyr + 6), fill=dotc)
        od.text((sb[0] + 56, cyr), name, font=font, fill=txtc, anchor="lm")
    img.alpha_composite(ov)
    cbox = (sb[2] + 1, y0 + tb_h + 4, x1, y1)
    content_fn(img, cbox)


KPIS = [
    ("MONTHLY REVENUE", "$11,600", "this month"),
    ("NET PROFIT", "$4,850", "after expenses"),
    ("OCCUPANCY RATE", "66%", "59 of 90 nights"),
    ("AVG NIGHTLY RATE", "$170", "portfolio ADR"),
    ("UPCOMING CHECK-INS", "3", "next 7 days"),
    ("UPCOMING CHECK-OUTS", "3", "next 7 days"),
    ("AVG LENGTH OF STAY", "3.9", "nights"),
    ("GUEST RATING", "4.8", "★ average"),
    ("CLEANING PENDING", "4", "turnovers"),
    ("MAINTENANCE DUE", "4", "open tasks"),
    ("LOW-STOCK ITEMS", "4", "reorder now"),
    ("BUSINESS HEALTH", "80%", "health score"),
]


def content_dashboard(img, cbox):
    x0, y0, x1, y1 = cbox; pad = 30
    d = ImageDraw.Draw(img); d.rectangle(cbox, fill=BG)
    d.text((x0 + pad, y0 + 22), "Executive Host Dashboard", font=fs(33), fill=PRIMARY, anchor="lt")
    d.text((x0 + pad, y0 + 64), "Summit Stays Co. · 3 properties  ·  your whole rental, automatically organized", font=fs(19, bold=False), fill=TEXT_MUTED, anchor="lt")
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
    d.text((gx, cy_top), "FINANCES · OCCUPANCY · SOURCES", font=fs(20), fill=ACCENT, anchor="lt")
    panels_y = cy_top + 34; panel_h = (y1 - panels_y - pad); pw = (gw - 3 * gap) / 4
    # expense breakdown donut
    px = gx
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Expense Breakdown", font=fs(17), fill=ACCENT, anchor="lt")
    donut(d, px + pw * 0.42, panels_y + panel_h * 0.52, min(panel_h * 0.28, pw * 0.27),
          [(47, PRIMARY), (14, ACCENT), (8, HIGHLIGHT), (6, SURFACE), (25, (170, 150, 120))], "$6.8K", "expenses")
    legend(d, px + pw * 0.04, panels_y + panel_h - 104, [(PRIMARY, "Mortgage 47%"), (ACCENT, "Cleaning 14%"), (HIGHLIGHT, "Taxes 8%")], 15, 30)
    # revenue by property (money bars)
    px = gx + (pw + gap)
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Revenue by Property", font=fs(17), fill=ACCENT, anchor="lt")
    hbars_money(img, d, (px + 20, panels_y + 56, px + pw - 16, panels_y + panel_h - 20),
                [("Pine Cabin", 1.0, "$5.2K"), ("Beachside", 0.75, "$3.9K"), ("Downtown Loft", 0.48, "$2.5K")])
    # occupancy by property
    px = gx + 2 * (pw + gap)
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Occupancy by Property", font=fs(17), fill=ACCENT, anchor="lt")
    hbars(img, d, (px + 20, panels_y + 56, px + pw - 16, panels_y + panel_h - 20),
          [("Pine Cabin", 0.80), ("Beachside", 0.63), ("Downtown Loft", 0.53)])
    # booking sources donut
    px = gx + 3 * (pw + gap)
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Booking Sources", font=fs(17), fill=ACCENT, anchor="lt")
    donut(d, px + pw * 0.42, panels_y + panel_h * 0.52, min(panel_h * 0.28, pw * 0.27),
          [(67, PRIMARY), (20, HIGHLIGHT), (13, ACCENT)], "15", "bookings")
    legend(d, px + pw * 0.04, panels_y + panel_h - 104, [(PRIMARY, "Airbnb 67%"), (HIGHLIGHT, "VRBO 20%"), (ACCENT, "Direct 13%")], 15, 30)


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
            if status_map is not None and ci == status_col:
                bg, fg = status_map.get(str(val), ((235, 230, 222), TEXT_MUTED))
                d.rounded_rectangle((hx - 68, ry + rh / 2 - 15, hx + 68, ry + rh / 2 + 15), radius=14, fill=bg)
                d.text((hx, ry + rh / 2), str(val), font=fs(14), fill=fg, anchor="mm")
            else:
                col = PRIMARY if ci == 0 else TEXT
                d.text((hx, ry + rh / 2), str(val), font=fs(16) if ci == 0 else fs(15, bold=False), fill=col, anchor=anc)
    if total_row:
        ry = ty + hdr_h + len(rows) * rh
        d.rectangle((tx0, ry, tx1, ry + rh), fill=SURFACE)
        for ci, val in enumerate(total_row):
            anc = "lm" if ci == 0 else "mm"
            hx = colx[ci] + (14 if ci == 0 else 0)
            if val != "":
                d.text((hx, ry + rh / 2), str(val), font=fs(16), fill=PRIMARY, anchor=anc)


def content_financial(img, cbox):
    x0, y0, x1, y1 = cbox; pad = 30
    d = ImageDraw.Draw(img); d.rectangle(cbox, fill=BG)
    d.text((x0 + pad, y0 + 20), "Financial Command Center", font=fs(31), fill=PRIMARY, anchor="lt")
    d.text((x0 + pad, y0 + 60), "This month's P&L — revenue, expenses, net profit & margin, computed live", font=fs(18, bold=False), fill=TEXT_MUTED, anchor="lt")
    strip = [("REVENUE", "$11,600"), ("EXPENSES", "$6,750"), ("NET PROFIT", "$4,850"), ("MARGIN", "42%")]
    sy = y0 + 92
    sw = (x1 - x0 - 2 * pad - 3 * 14) / 4
    for i, (lab, val) in enumerate(strip):
        sx = x0 + pad + i * (sw + 14)
        d.rounded_rectangle((sx, sy, sx + sw, sy + 90), radius=10, fill=WHITE, outline=GRID, width=2)
        d.rounded_rectangle((sx + 12, sy, sx + sw - 12, sy + 5), radius=2, fill=GOLD_LT)
        d.text((sx + sw / 2, sy + 30), lab, font=fs(15), fill=ACCENT, anchor="mm")
        col = PRIMARY if lab != "NET PROFIT" else (32, 120, 96)
        d.text((sx + sw / 2, sy + 64), val, font=fserif(28), fill=col, anchor="mm")
    rows = [
        ("Mortgage", 3200), ("Utilities", 420), ("Internet", 95), ("Insurance", 180),
        ("Cleaning", 960), ("Supplies", 340), ("Repairs", 250), ("Property Taxes", 520),
        ("HOA Fees", 150), ("Software", 85), ("Marketing", 120), ("Platform Fees", 348),
        ("Miscellaneous", 82),
    ]
    total = 6750
    ty = sy + 112
    tx0, tx1 = x0 + pad, x1 - pad
    headers = ["EXPENSE CATEGORY", "AMOUNT", "% OF SPEND"]
    colx = [tx0, tx0 + (tx1 - tx0) * 0.46, tx0 + (tx1 - tx0) * 0.62]
    hdr_h = 40
    grad_round(img, (tx0, ty, tx1, ty + hdr_h), 8, PRIMARY_LT, PRIMARY_DK)
    for i, h in enumerate(headers):
        anc = "lm" if i == 0 else ("mm" if i == 1 else "lm")
        hx = colx[i] + (14 if i != 1 else (tx1 - tx0) * 0.07)
        d.text((hx, ty + hdr_h / 2), h, font=fs(15), fill=WHITE, anchor=anc)
    rh = (y1 - pad - (ty + hdr_h)) / (len(rows) + 1)
    barw = (tx1 - tx0) * 0.30
    for i, (cat, amt) in enumerate(rows):
        ry = ty + hdr_h + i * rh
        if i % 2:
            d.rectangle((tx0, ry, tx1, ry + rh), fill=ROW_ALT)
        d.text((tx0 + 14, ry + rh / 2), cat, font=fs(15, bold=False), fill=TEXT, anchor="lm")
        d.text((colx[1] + (tx1 - tx0) * 0.07, ry + rh / 2), f"${amt:,}", font=fs(15), fill=PRIMARY, anchor="mm")
        frac = amt / 3200
        tbx0 = colx[2] + 14
        d.rounded_rectangle((tbx0, ry + rh / 2 - 9, tbx0 + barw, ry + rh / 2 + 9), radius=8, fill=(236, 230, 220))
        grad_round(img, (tbx0, ry + rh / 2 - 9, tbx0 + barw * frac, ry + rh / 2 + 9), 8, PRIMARY_LT, PRIMARY_DK)
        d.text((tbx0 + barw + 12, ry + rh / 2), f"{amt/total*100:.0f}%", font=fs(13), fill=ACCENT, anchor="lm")
    ry = ty + hdr_h + len(rows) * rh
    d.rectangle((tx0, ry, tx1, ry + rh), fill=SURFACE)
    d.text((tx0 + 14, ry + rh / 2), "TOTAL EXPENSES", font=fs(17), fill=PRIMARY, anchor="lm")
    d.text((colx[1] + (tx1 - tx0) * 0.07, ry + rh / 2), "$6,750", font=fs(16), fill=PRIMARY, anchor="mm")
    d.text((colx[2] + 14, ry + rh / 2), "58% of revenue", font=fs(14), fill=PRIMARY, anchor="lm")


def content_calendar(img, cbox):
    bookings = [
        ("Sarah Mitchell", "Airbnb", -22, 5, "The Pine Cabin", "Completed"),
        ("James Carter", "Airbnb", -18, 3, "Downtown Loft", "Completed"),
        ("The Nguyen Family", "VRBO", -14, 6, "Beachside Bungalow", "Completed"),
        ("Emily Rodriguez", "Airbnb", -9, 4, "The Pine Cabin", "Completed"),
        ("David Kim", "Airbnb", -2, 4, "The Pine Cabin", "In Progress"),
        ("Priya Patel", "VRBO", -1, 5, "Beachside Bungalow", "In Progress"),
        ("Tom & Lisa Hayes", "Airbnb", 2, 3, "Downtown Loft", "Confirmed"),
        ("Grace Liu", "Airbnb", 4, 6, "The Pine Cabin", "Confirmed"),
        ("Robert Fox", "Direct", 6, 2, "Beachside Bungalow", "Confirmed"),
        ("Marcus Webb", "VRBO", 14, 5, "The Pine Cabin", "Confirmed"),
    ]
    today = dt.date.today()
    rows = []
    for g, src, off, nights, prop, status in bookings:
        ci = (today + dt.timedelta(days=off)).strftime("%b %d")
        rows.append((g, src, ci, str(nights), prop, status))
    _table(img, cbox, "Master Booking Calendar",
           "Every stay across every property — occupancy & check-ins calculate automatically",
           ["GUEST", "SOURCE", "CHECK-IN", "NIGHTS", "PROPERTY", "STATUS"],
           [0.0, 0.24, 0.37, 0.50, 0.60, 0.86], rows,
           status_col=5, status_map={"Completed": (SURFACE, (110, 88, 58)),
                                      "In Progress": (WARN_BG, ACCENT),
                                      "Confirmed": (MINT_BG, PRIMARY)})


def content_reservations(img, cbox):
    rows = [
        ("RES-1041", "Sarah Mitchell", "The Pine Cabin", "5", "$1,150", "$1,143"),
        ("RES-1042", "James Carter", "Downtown Loft", "3", "$480", "$502"),
        ("RES-1043", "The Nguyen Family", "Beachside Bungalow", "6", "$1,260", "$1,211"),
        ("RES-1044", "Emily Rodriguez", "The Pine Cabin", "4", "$940", "$956"),
        ("RES-1045", "Michael Brooks", "Downtown Loft", "2", "$320", "$369"),
        ("RES-1046", "Olivia Bennett", "Beachside Bungalow", "3", "$660", "$737"),
        ("RES-1047", "David Kim", "The Pine Cabin", "4", "$940", "$916"),
        ("RES-1048", "Priya Patel", "Beachside Bungalow", "5", "$1,150", "$1,173"),
        ("RES-1049", "Tom & Lisa Hayes", "Downtown Loft", "3", "$480", "$502"),
        ("RES-1050", "Grace Liu", "The Pine Cabin", "6", "$1,440", "$1,401"),
    ]
    _table(img, cbox, "Reservation Manager",
           "Per-booking money, decoded — gross, fees, taxes & net revenue for every stay",
           ["RES ID", "GUEST", "PROPERTY", "NIGHTS", "GROSS", "NET REVENUE"],
           [0.0, 0.17, 0.40, 0.60, 0.72, 0.86], rows)


# ---------- renders ----------

def render_hero(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=640)
    d = ImageDraw.Draw(img)
    house_crest(img, SIZE // 2, 132, r=56)
    pill(img, SIZE // 2, 256, "THE ULTIMATE AIRBNB BUSINESS MANAGEMENT SYSTEM", font=fs(26), pad_x=42, pad_y=20)
    wordmark(img, SIZE // 2, 400, "AIRBNB HOST COMMAND CENTER", 100, max_w=1840)
    gold_divider(img, SIZE // 2, 500, width=520)
    tc(d, (SIZE // 2, 550), "Reservations, finances, pricing, cleaning & reviews — run your rental like a business.",
       fs(24, bold=False), (224, 213, 190))
    chips = [("19", "POWERFUL TABS"), ("AUTO", "PROFIT + OCCUPANCY"), ("2-in-1", "EXCEL + SHEETS")]
    cw = 420
    total = len(chips) * cw + (len(chips) - 1) * 32
    startx = (SIZE - total) // 2 + cw // 2
    for i, (b, s) in enumerate(chips):
        stat_chip(img, startx + i * (cw + 32), 704, b, s, w=cw)
    app_window(img, (70, 800, SIZE - 70, 1900), 0, content_dashboard)
    pill(img, SIZE // 2, SIZE - 52, "19 SHEETS · INSTANT DOWNLOAD · EXCEL + GOOGLE SHEETS",
         font=fs(33), pad_x=50, pad_y=24, star=True, grad=(PRIMARY_LT, PRIMARY_DK))
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_inside(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=400)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 120, "EVERYTHING INSIDE", font=fs(38), pad_x=54, pad_y=22)
    tc(d, (SIZE // 2, 238), "19 Powerful, Connected Tabs", fserif(58), WHITE)
    gold_divider(img, SIZE // 2, 308, width=520)
    tc(d, (SIZE // 2, 352), "Not a booking spreadsheet — a complete vacation-rental operating system",
       fs(24, bold=False), (226, 214, 190))
    cards = [
        ("Host Dashboard", "12 KPIs + live charts"), ("Property Profile", "listing facts & codes"),
        ("Booking Calendar", "auto occupancy"), ("Reservations", "per-booking net revenue"),
        ("Financial Center", "live P&L"), ("Pricing Strategy", "scenario modeler"),
        ("Guest CRM", "repeat-guest engine"), ("Cleaning Center", "turnover scheduling"),
        ("Maintenance", "priority & cost"), ("Inventory", "low-stock alerts"),
        ("Review Tracker", "scored by category"), ("Message Library", "reusable templates"),
        ("Supplier Directory", "your trusted team"), ("Tax Prep", "deduction-ready"),
        ("Multi-Property", "compare listings"), ("Goal Planner", "targets & milestones"),
        ("Improvements", "photo & upgrade plan"), ("Analytics", "business health score"),
        ("Settings", "make it yours"),
    ]
    cols = 4
    margin = 90
    gx, gy = 22, 22
    cw = (SIZE - 2 * margin - (cols - 1) * gx) // cols
    top = 440
    rows_n = 5
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
        od.text((x + 82, cyc - 17), title, font=fs(20), fill=PRIMARY, anchor="lm")
        od.text((x + 82, cyc + 22), sub, font=fs(15, bold=False), fill=TEXT_MUTED, anchor="lm")
        img.alpha_composite(ov)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_dashboard(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=360)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 116, "EXECUTIVE HOST DASHBOARD", font=fs(34), pad_x=50, pad_y=22)
    tc(d, (SIZE // 2, 232), "Your Whole Rental, At A Glance", fserif(52), WHITE)
    tc(d, (SIZE // 2, 300), "12 live KPIs + finances, occupancy, sources & business-health charts — all auto-updating",
       fs(23, bold=False), (226, 214, 190))
    app_window(img, (70, 400, SIZE - 70, SIZE - 70), 0, content_dashboard)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_financial(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=320)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 112, "FINANCIAL COMMAND CENTER", font=fs(34), pad_x=52, pad_y=22)
    tc(d, (SIZE // 2, 230), "Know Your Real Profit — Every Month", fserif(48), WHITE)
    tc(d, (SIZE // 2, 290), "Live P&L · 13 expense categories · net profit, margin & cash flow",
       fs(23, bold=False), (226, 214, 190))
    app_window(img, (80, 360, SIZE - 80, SIZE - 70), 4, content_financial)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_manage(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=300)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 110, "BOOKINGS & REVENUE", font=fs(36), pad_x=52, pad_y=22)
    tc(d, (SIZE // 2, 224), "Every Stay Booked, Every Dollar Tracked", fserif(42), WHITE)
    app_window(img, (60, 330, SIZE - 60, 1150), 2, content_calendar)
    app_window(img, (60, 1180, SIZE - 60, SIZE - 60), 3, content_reservations)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_mobile(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=400)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 130, "WORKS EVERYWHERE", font=fs(38), pad_x=54, pad_y=22)
    tc(d, (SIZE // 2, 250), "Excel · Google Sheets · Mobile", fserif(56), WHITE)
    tc(d, (SIZE // 2, 320), "Approve a booking or check tonight's turnover from anywhere — your whole business in your pocket",
       fs(22, bold=False), (226, 214, 190))
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
    od.text(((sx0 + sx1) // 2, sy0 + 56), "Summit Stays", font=fserif(34), fill=GOLD_HI, anchor="mm")
    y = sy0 + 150
    cards = [("MONTHLY REVENUE", "$11,600", PRIMARY), ("OCCUPANCY RATE", "66%", PRIMARY),
             ("GUEST RATING", "4.8 / 5", ACCENT), ("BUSINESS HEALTH", "80%", PRIMARY)]
    for lab, val, col in cards:
        cb = (sx0 + 30, y, sx1 - 30, y + 135)
        od.rounded_rectangle(cb, radius=18, fill=WHITE, outline=GOLD_LT, width=2)
        od.rounded_rectangle((cb[0] + 20, y, cb[2] - 20, y + 5), radius=2, fill=GOLD_LT)
        od.text((cb[0] + 26, y + 34), lab, font=fs(22), fill=ACCENT, anchor="lt")
        vf = fit_font(od, val, sx1 - sx0 - 110, 46)
        od.text((cb[0] + 26, y + 94), val, font=vf, fill=col, anchor="lm")
        y += 152
    od.text((sx0 + 40, y + 16), "TODAY'S TASKS", font=fs(22), fill=ACCENT, anchor="lt")
    y += 52
    for lab, state in [("Approve Grace Liu booking", False), ("Restock coffee — Pine Cabin", False),
                       ("Message David Kim (checkout)", True), ("Confirm Friday cleaner", True)]:
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
        ("04_financial.png", render_financial),
        ("05_manage.png", render_manage),
        ("06_mobile.png", render_mobile),
    ]
    for name, fn in targets:
        fn(os.path.join(out_dir, name))
        print(f"  ✓ {name}")
    print(f"Wrote {len(targets)} images to {out_dir}")


if __name__ == "__main__":
    main()

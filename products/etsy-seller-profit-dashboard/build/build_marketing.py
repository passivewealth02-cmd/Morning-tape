"""Marketing image set for Etsy Seller Profit Dashboard™ (6 images, 2000x2000).

Dense "finance-app screenshots" that mirror the real workbook: a left sidebar
of all 14 tabs, the REAL computed numbers from the sample data, and fully
populated tables/charts (the order engine, fee breakdown, product ranking).

  01_hero.png            - branded hero + live seller dashboard
  02_inside.png          - "everything inside — 14 powerful tabs"
  03_dashboard.png       - full seller executive dashboard
  04_orders.png          - the order tracker (auto fees & profit)
  05_fees_products.png   - fee breakdown + product ranking
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

TABS = ["Dashboard", "Orders", "Product Calc", "Fees", "Ads", "Expenses",
        "Library", "Monthly", "Insights", "Goals", "Cash Flow", "Tax Prep",
        "Strategy", "Settings"]


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


def profit_crest(c, cx, cy, r=60, glow=True):
    if glow:
        radial_glow(c, cx, cy, int(r * 2.1), GOLD_HI, 90)
    grad_round(c, (cx - r, cy - r, cx + r, cy + r), 22, PRIMARY_LT, PRIMARY_DK, outline=GOLD_LT, width=4)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    d.rounded_rectangle((cx - r + 10, cy - r + 10, cx + r - 10, cy + r - 10), radius=16, outline=GOLD_HI, width=2)
    # ascending bars
    bx = cx - r * 0.5
    bw = r * 0.22
    heights = [0.35, 0.55, 0.8]
    basey = cy + r * 0.45
    for i, hh in enumerate(heights):
        x = bx + i * (bw + r * 0.09)
        d.rounded_rectangle((x, basey - r * hh, x + bw, basey), radius=3, fill=GOLD_HI)
    # upward arrow line
    p0 = (cx - r * 0.5, cy + r * 0.18)
    p1 = (cx + r * 0.5, cy - r * 0.5)
    d.line([p0, p1], fill=HIGHLIGHT, width=6)
    d.polygon([(p1[0] + 2, p1[1] - 2), (p1[0] - r * 0.22, p1[1] + r * 0.02),
               (p1[0] + r * 0.02, p1[1] + r * 0.22)], fill=HIGHLIGHT)
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
        d.text((cx, cy - (10 if center_sub else 0)), center_top, font=fserif(int(r * 0.32)), fill=PRIMARY, anchor="mm")
    if center_sub:
        d.text((cx, cy + int(r * 0.26)), center_sub, font=fs(int(r * 0.13)), fill=TEXT_MUTED, anchor="mm")


def legend(d, x, y, items, fsz=20, gap=40):
    for i, (col, lab) in enumerate(items):
        yy = y + i * gap
        d.rounded_rectangle((x, yy, x + 24, yy + 24), radius=5, fill=col)
        d.text((x + 36, yy + 12), lab, font=fs(fsz, bold=False), fill=TEXT, anchor="lm")


def vbars(img, d, box, items, maxv=1.0, suffix=""):
    x0, y0, x1, y1 = box
    n = len(items)
    bw = (x1 - x0) / n
    base = y1 - 36
    for i, (lab, val) in enumerate(items):
        cx = x0 + bw * (i + 0.5)
        h = (val / maxv) * (base - y0 - 14)
        grad_round(img, (cx - bw * 0.28, base - h, cx + bw * 0.28, base), 8, HIGHLIGHT, (70, 200, 165))
        vt = f"${int(val)}" if suffix == "$" else (f"{val:g}{suffix}")
        d.text((cx, base - h - 20), vt, font=fs(17), fill=PRIMARY, anchor="mm")
        d.text((cx, base + 18), lab, font=fs(15, bold=False), fill=TEXT_MUTED, anchor="mm")


def mini_lines(d, box, series, colors):
    x0, y0, x1, y1 = box
    allv = [v for s in series for v in s]
    mx = max(allv) if allv else 1
    n = len(series[0])
    for si, s in enumerate(series):
        pts = []
        for i, v in enumerate(s):
            px = x0 + (x1 - x0) * (i / (n - 1))
            py = y1 - (y1 - y0) * (v / mx)
            pts.append((px, py))
        d.line(pts, fill=colors[si], width=5, joint="curve")
        for p in pts:
            d.ellipse((p[0] - 5, p[1] - 5, p[0] + 5, p[1] + 5), fill=colors[si])


def fit_font(d, text, max_w, start, serif=True):
    s = start
    f = fserif(s) if serif else fs(s)
    while s > 12 and d.textlength(text, font=f) > max_w:
        s -= 1
        f = fserif(s) if serif else fs(s)
    return f


def app_window(img, box, active_idx, content_fn, file_label="Etsy_Seller_Profit_Dashboard.xlsx — Willow & Oak Co."):
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
    od.text((bx, sb[1] + 30), "ETSY PROFIT", font=fs(19), fill=GOLD_HI, anchor="lt")
    od.text((bx, sb[1] + 56), "14-tab CFO system", font=fs(15, bold=False), fill=(170, 200, 192), anchor="lt")
    od.line((sb[0] + 20, sb[1] + 88, sb[2] - 16, sb[1] + 88), fill=(255, 255, 255, 40), width=1)
    list_top = sb[1] + 104
    rowh = (y1 - 24 - list_top) / len(TABS)
    palette = [HIGHLIGHT, GOLD_HI, SURFACE, (150, 200, 190)]
    for i, name in enumerate(TABS):
        ry = list_top + i * rowh
        if i == active_idx:
            od.rounded_rectangle((sb[0] + 12, ry + 2, sb[2] - 10, ry + rowh - 2), radius=8, fill=(255, 255, 255, 235))
            od.rounded_rectangle((sb[0] + 12, ry + 2, sb[0] + 19, ry + rowh - 2), radius=3, fill=GOLD_HI)
            dotc = PRIMARY; txtc = PRIMARY; font = fs(19)
        else:
            dotc = palette[i % len(palette)]; txtc = (214, 226, 222); font = fs(18, bold=False)
        cyr = ry + rowh / 2
        od.ellipse((sb[0] + 30, cyr - 6, sb[0] + 42, cyr + 6), fill=dotc)
        od.text((sb[0] + 56, cyr), name, font=font, fill=txtc, anchor="lm")
    img.alpha_composite(ov)
    cbox = (sb[2] + 1, y0 + tb_h + 4, x1, y1)
    content_fn(img, cbox)


KPIS = [
    ("TOTAL REVENUE", "$761", "this month"),
    ("NET PROFIT", "$530", "after everything"),
    ("ETSY FEES PAID", "$127", "17% of revenue"),
    ("AD SPEND", "$55", "Etsy Ads"),
    ("REFUNDS", "$20", "2 orders"),
    ("PROFIT MARGIN", "70%", "your TRUE margin"),
    ("AVG ORDER VALUE", "$17.70", "per sale"),
    ("ORDERS", "43", "this month"),
    ("BEST SELLER", "Mega Bundle", "$232 revenue"),
    ("NEEDS WORK", "Sticker Pack", "$15 profit"),
]


def content_dashboard(img, cbox):
    x0, y0, x1, y1 = cbox
    pad = 30
    d = ImageDraw.Draw(img)
    d.rectangle(cbox, fill=BG)
    d.text((x0 + pad, y0 + 22), "Seller Executive Dashboard", font=fs(33), fill=PRIMARY, anchor="lt")
    d.text((x0 + pad, y0 + 64), "Willow & Oak Co.  ·  revenue is vanity — profit is sanity", font=fs(19, bold=False), fill=TEXT_MUTED, anchor="lt")
    d.rounded_rectangle((x1 - pad - 150, y0 + 26, x1 - pad, y0 + 62), radius=18, fill=MINT_BG)
    d.text((x1 - pad - 75, y0 + 44), "● live", font=fs(18), fill=PRIMARY, anchor="mm")
    gx = x0 + pad; gy = y0 + 98
    gw = (x1 - x0 - 2 * pad); gap = 15
    kw = (gw - 4 * gap) / 5; kh = 122
    for i, (lab, val, sub) in enumerate(KPIS):
        r, ci = divmod(i, 5)
        kx = gx + ci * (kw + gap); ky = gy + r * (kh + gap)
        d.rounded_rectangle((kx, ky, kx + kw, ky + kh), radius=12, fill=WHITE, outline=GRID, width=2)
        d.rounded_rectangle((kx + 14, ky, kx + kw - 14, ky + 5), radius=2, fill=GOLD_LT)
        d.text((kx + 16, ky + 18), lab, font=fs(13), fill=ACCENT, anchor="lt")
        vf = fit_font(d, val, kw - 32, 32)
        d.text((kx + 16, ky + 66), val, font=vf, fill=PRIMARY, anchor="lm")
        d.text((kx + 16, ky + 100), sub, font=fs(13, bold=False), fill=TEXT_MUTED, anchor="lm")
    cy_top = gy + 2 * (kh + gap) + 18
    d.text((gx, cy_top), "PROFIT · FEES · PRODUCTS · ADS", font=fs(20), fill=ACCENT, anchor="lt")
    panels_y = cy_top + 34
    panel_h = (y1 - panels_y - pad)
    pw = (gw - 3 * gap) / 4
    # revenue vs profit line
    px = gx
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Revenue vs Profit", font=fs(17), fill=ACCENT, anchor="lt")
    mini_lines(d, (px + 26, panels_y + 60, px + pw - 20, panels_y + panel_h - 70),
               [[210, 340, 480, 610, 720, 740], [118, 221, 312, 430, 512, 521]], [PRIMARY, HIGHLIGHT])
    legend(d, px + pw * 0.06, panels_y + panel_h - 56, [(PRIMARY, "Revenue"), (HIGHLIGHT, "Profit")], 15, 28)
    # fees donut
    px = gx + (pw + gap)
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Etsy Fees", font=fs(17), fill=ACCENT, anchor="lt")
    donut(d, px + pw * 0.42, panels_y + panel_h * 0.52, min(panel_h * 0.28, pw * 0.27),
          [(39, PRIMARY), (27, ACCENT), (26, (170, 150, 120)), (8, HIGHLIGHT)], "$127", "in fees")
    legend(d, px + pw * 0.04, panels_y + panel_h - 104, [(PRIMARY, "Txn 39%"), (ACCENT, "Proc 27%"), ((170, 150, 120), "Offsite 26%")], 15, 30)
    # profit by product
    px = gx + 2 * (pw + gap)
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Profit by Product", font=fs(17), fill=ACCENT, anchor="lt")
    vbars(img, d, (px + 16, panels_y + 50, px + pw - 12, panels_y + panel_h - 6),
          [("Mega", 189), ("Wed", 143), ("Bud", 137), ("Res", 79), ("IG", 53), ("Stk", 15)], maxv=189, suffix="$")
    # ads ROAS stat
    px = gx + 3 * (pw + gap)
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Ad Efficiency", font=fs(17), fill=ACCENT, anchor="lt")
    ccx, ccy = px + pw / 2, panels_y + panel_h * 0.46
    d.text((ccx, ccy), "7.3x", font=fserif(int(pw * 0.28)), fill=PRIMARY, anchor="mm")
    d.text((ccx, ccy + panel_h * 0.20), "blended ROAS", font=fs(18), fill=ACCENT, anchor="mm")
    d.text((ccx, panel_h + panels_y - 40), "$399 revenue  ·  $55 spent", font=fs(16, bold=False), fill=TEXT_MUTED, anchor="mm")


def _table(img, cbox, title, subtitle, headers, colf, rows, active_first_lm=True,
           total_row=None, status_col=None, status_map=None, hdr_top=104):
    x0, y0, x1, y1 = cbox
    pad = 30
    d = ImageDraw.Draw(img)
    d.rectangle(cbox, fill=BG)
    d.text((x0 + pad, y0 + 22), title, font=fs(32), fill=PRIMARY, anchor="lt")
    d.text((x0 + pad, y0 + 62), subtitle, font=fs(18, bold=False), fill=TEXT_MUTED, anchor="lt")
    tx0, tx1 = x0 + pad, x1 - pad
    ty = y0 + hdr_top
    n = len(headers)
    colx = [tx0 + (tx1 - tx0) * f for f in colf]
    hdr_h = 44
    grad_round(img, (tx0, ty, tx1, ty + hdr_h), 8, PRIMARY_LT, PRIMARY_DK)
    for i, h in enumerate(headers):
        anc = "lm" if i == 0 else "mm"
        hx = colx[i] + (14 if i == 0 else 0)
        d.text((hx, ty + hdr_h / 2), h, font=fs(15), fill=WHITE, anchor=anc)
    nrows = len(rows) + (1 if total_row else 0)
    rh = (y1 - pad - (ty + hdr_h)) / nrows
    for i, row in enumerate(rows):
        ry = ty + hdr_h + i * rh
        refunded = status_col is not None and str(row[status_col]) in ("Yes", "Refund")
        if refunded:
            d.rectangle((tx0, ry, tx1, ry + rh), fill=RED_BG)
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
                col = PRIMARY if (ci == 0 or (isinstance(val, str) and val.startswith("$") and ci >= n - 2)) else TEXT
                fnt = fs(17) if ci == 0 else fs(16, bold=(ci >= n - 1))
                d.text((hx, ry + rh / 2), str(val), font=fnt, fill=col, anchor=anc)
    if total_row:
        ry = ty + hdr_h + len(rows) * rh
        d.rectangle((tx0, ry, tx1, ry + rh), fill=SURFACE)
        for ci, val in enumerate(total_row):
            anc = "lm" if ci == 0 else "mm"
            hx = colx[ci] + (14 if ci == 0 else 0)
            if val != "":
                d.text((hx, ry + rh / 2), str(val), font=fs(17), fill=PRIMARY, anchor=anc)


def content_orders(img, cbox):
    rows = [
        ("#2401", "Budget Planner", "2", "$24", "$6.53", "No", "$17.47"),
        ("#2402", "Wedding Suite", "1", "$18", "$2.16", "No", "$15.84"),
        ("#2403", "Resume Bundle", "1", "$14", "$1.78", "No", "$12.22"),
        ("#2404", "IG Templates", "1", "$9", "$1.30", "No", "$7.70"),
        ("#2405", "Mega Bundle", "1", "$29", "$7.55", "No", "$21.45"),
        ("#2406", "Sticker Pack", "1", "$6", "$1.02", "Yes", "$0.00"),
        ("#2407", "Budget Planner", "1", "$12", "$1.59", "No", "$10.41"),
        ("#2408", "Wedding Suite", "2", "$36", "$4.07", "No", "$31.93"),
        ("#2409", "Resume Bundle", "1", "$14", "$3.88", "No", "$10.12"),
        ("#2410", "IG Templates", "1", "$9", "$1.30", "No", "$7.70"),
        ("#2411", "Mega Bundle", "1", "$29", "$3.21", "No", "$25.80"),
        ("#2412", "Sticker Pack", "1", "$6", "$1.02", "No", "$4.98"),
    ]
    _table(img, cbox, "Order Tracker",
           "Type in the sale — Etsy fees & net profit calculate themselves (refunds flag red)",
           ["ORDER", "PRODUCT", "QTY", "SALE", "ETSY FEES", "REFUND?", "NET PROFIT"],
           [0.0, 0.13, 0.40, 0.50, 0.62, 0.78, 0.90],
           rows, status_col=5, status_map={"Yes": (RED_BG, DANGER), "No": (MINT_BG, PRIMARY)},
           total_row=("TOTALS", "43 orders", "50", "$761", "$127", "", "$617"))


def content_fees(img, cbox):
    rows = [
        ("Transaction Fees", "$49.47", "6.5%"),
        ("Processing Fees", "$33.58", "4.4%"),
        ("Offsite Ads Fees", "$34.05", "4.5%"),
        ("Listing Fees", "$10.00", "1.3%"),
        ("Etsy Ads Spend", "$55.00", "7.2%"),
        ("Subscription", "$0.00", "0.0%"),
    ]
    x0, y0, x1, y1 = cbox
    pad = 30
    d = ImageDraw.Draw(img)
    d.rectangle(cbox, fill=BG)
    d.text((x0 + pad, y0 + 22), "Etsy Fees Breakdown", font=fs(32), fill=PRIMARY, anchor="lt")
    d.text((x0 + pad, y0 + 62), "Exactly what Etsy takes — and the % of every dollar you keep", font=fs(18, bold=False), fill=TEXT_MUTED, anchor="lt")
    # table left
    tx0 = x0 + pad
    tw = (x1 - x0) * 0.52
    tx1 = tx0 + tw
    ty = y0 + 104
    headers = ["FEE TYPE", "AMOUNT", "% REV"]
    colx = [tx0, tx0 + tw * 0.58, tx0 + tw * 0.82]
    hdr_h = 44
    grad_round(img, (tx0, ty, tx1, ty + hdr_h), 8, PRIMARY_LT, PRIMARY_DK)
    for i, h in enumerate(headers):
        d.text((colx[i] + (14 if i == 0 else 0), ty + hdr_h / 2), h, font=fs(15), fill=WHITE, anchor="lm" if i == 0 else "mm")
    rh = (y1 - pad - (ty + hdr_h)) / (len(rows) + 1)
    for i, (lab, amt, pct) in enumerate(rows):
        ry = ty + hdr_h + i * rh
        if i % 2:
            d.rectangle((tx0, ry, tx1, ry + rh), fill=ROW_ALT)
        d.text((colx[0] + 14, ry + rh / 2), lab, font=fs(17, bold=False), fill=TEXT, anchor="lm")
        d.text((colx[1], ry + rh / 2), amt, font=fs(17), fill=PRIMARY, anchor="mm")
        d.text((colx[2], ry + rh / 2), pct, font=fs(16, bold=False), fill=ACCENT, anchor="mm")
    ry = ty + hdr_h + len(rows) * rh
    d.rectangle((tx0, ry, tx1, ry + rh), fill=SURFACE)
    d.text((colx[0] + 14, ry + rh / 2), "TOTAL COST OF SELLING", font=fs(16), fill=PRIMARY, anchor="lm")
    d.text((colx[1], ry + rh / 2), "$182", font=fs(17), fill=PRIMARY, anchor="mm")
    d.text((colx[2], ry + rh / 2), "24%", font=fs(16), fill=PRIMARY, anchor="mm")
    # keep ring right
    rcx = tx1 + (x1 - pad - tx1) / 2
    rcy = (ty + y1 - pad) / 2
    rr = min((x1 - pad - tx1) * 0.34, (y1 - ty) * 0.30)
    radial_glow(img, int(rcx), int(rcy), int(rr * 1.4), HIGHLIGHT, 55)
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    od.ellipse((rcx - rr, rcy - rr, rcx + rr, rcy + rr), fill=WHITE, outline=(232, 226, 214), width=3)
    od.pieslice((rcx - rr, rcy - rr, rcx + rr, rcy + rr), -90, -90 + 0.69 * 360, fill=HIGHLIGHT)
    od.ellipse((rcx - rr + 48, rcy - rr + 48, rcx + rr - 48, rcy + rr - 48), fill=WHITE)
    od.text((rcx, rcy - 12), "70%", font=fserif(int(rr * 0.5)), fill=PRIMARY, anchor="mm")
    od.text((rcx, rcy + rr * 0.4), "YOU KEEP", font=fs(int(rr * 0.16)), fill=ACCENT, anchor="mm")
    od.text((rcx, rcy + rr + 34), "$530 net profit of $761 revenue", font=fs(18, bold=False), fill=TEXT_MUTED, anchor="mm")
    img.alpha_composite(ov)


def content_library(img, cbox):
    rows = [
        ("Mega Bundle", "Printable", "8", "$232", "$189", "81%"),
        ("Wedding Suite", "Printable", "9", "$162", "$143", "88%"),
        ("Budget Planner", "Printable", "14", "$168", "$137", "82%"),
        ("Resume Bundle", "Template", "8", "$112", "$79", "71%"),
        ("IG Templates", "Digital", "7", "$63", "$53", "84%"),
        ("Sticker Pack", "Digital", "4", "$24", "$15", "62%"),
    ]
    _table(img, cbox, "Product Library",
           "Every listing ranked by profit — scale the winners, fix or cut the rest",
           ["PRODUCT", "CATEGORY", "UNITS", "REVENUE", "PROFIT", "MARGIN"],
           [0.0, 0.30, 0.52, 0.64, 0.78, 0.90],
           rows)


# ---------- renders ----------

def render_hero(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=640)
    d = ImageDraw.Draw(img)
    profit_crest(img, SIZE // 2, 132, r=56)
    pill(img, SIZE // 2, 256, "THE ULTIMATE ETSY FINANCE & PROFIT SYSTEM · 2026", font=fs(27), pad_x=42, pad_y=20)
    wordmark(img, SIZE // 2, 372, "ETSY SELLER", 108, max_w=1500)
    wordmark(img, SIZE // 2, 472, "PROFIT DASHBOARD", 100, max_w=1620)
    gold_divider(img, SIZE // 2, 542, width=520)
    tc(d, (SIZE // 2, 590), "Stop guessing. See exactly what you keep after every Etsy fee, ad & refund.",
       fs(26, bold=False), (224, 213, 190))
    chips = [("14", "POWERFUL TABS"), ("AUTO", "PROFIT PER ORDER"), ("2-in-1", "EXCEL + SHEETS")]
    cw = 420
    total = len(chips) * cw + (len(chips) - 1) * 32
    startx = (SIZE - total) // 2 + cw // 2
    for i, (b, s) in enumerate(chips):
        stat_chip(img, startx + i * (cw + 32), 706, b, s, w=cw)
    app_window(img, (70, 800, SIZE - 70, 1900), 0, content_dashboard)
    pill(img, SIZE // 2, SIZE - 52, "14 SHEETS · INSTANT DOWNLOAD · EXCEL + GOOGLE SHEETS",
         font=fs(33), pad_x=50, pad_y=24, star=True, grad=(PRIMARY_LT, PRIMARY_DK))
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_inside(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=400)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 120, "EVERYTHING INSIDE", font=fs(38), pad_x=54, pad_y=22)
    tc(d, (SIZE // 2, 238), "14 Powerful, Connected Tabs", fserif(58), WHITE)
    gold_divider(img, SIZE // 2, 308, width=520)
    tc(d, (SIZE // 2, 352), "Not a tracker — a mini Etsy CFO system that turns shop data into profit intelligence",
       fs(24, bold=False), (226, 214, 190))
    cards = [
        ("Seller Dashboard", "10 KPIs + 4 live charts"), ("Order Tracker", "auto fees & net profit"),
        ("Product Calculator", "true margin per listing"), ("Etsy Fees Breakdown", "what Etsy really takes"),
        ("Ads Performance", "ROAS & profit after ads"), ("Expense Tracker", "tools, courses & more"),
        ("Product Library", "every listing ranked"), ("Monthly Snapshot", "revenue vs profit trend"),
        ("Customer Insights", "AOV & refund rate"), ("Goal Tracker", "live progress bars"),
        ("Cash Flow", "money in vs out"), ("Tax Prep", "set-aside & reserve"),
        ("Strategy Analyzer", "scale / fix / cut"), ("Settings", "your fee rates & goals"),
    ]
    cols = 4
    margin = 90
    gx, gy = 22, 22
    cw = (SIZE - 2 * margin - (cols - 1) * gx) // cols
    top = 450
    rows_n = 4
    ch = (SIZE - top - 70 - (rows_n - 1) * gy) // rows_n
    for i, (title, sub) in enumerate(cards):
        r, ccol = divmod(i, cols)
        x = margin + ccol * (cw + gx); y = top + r * (ch + gy)
        shadow(img, (x, y, x + cw, y + ch), 14, 14, 45, 10)
        ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
        od = ImageDraw.Draw(ov)
        od.rounded_rectangle((x, y, x + cw, y + ch), radius=14, fill=WHITE, outline=(232, 224, 208), width=2)
        od.rounded_rectangle((x, y, x + 7, y + ch), radius=3, fill=GOLD_LT)
        od.rectangle((x + 3, y, x + 7, y + ch), fill=GOLD_LT)
        cyc = y + ch // 2
        bx = x + 46
        od.ellipse((bx - 26, cyc - 26, bx + 26, cyc + 26), fill=PRIMARY)
        od.text((bx, cyc), str(i + 1), font=fs(23), fill=GOLD_HI, anchor="mm")
        od.text((x + 92, cyc - 18), title, font=fs(23), fill=PRIMARY, anchor="lm")
        od.text((x + 92, cyc + 22), sub, font=fs(17, bold=False), fill=TEXT_MUTED, anchor="lm")
        img.alpha_composite(ov)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_dashboard(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=360)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 116, "SELLER EXECUTIVE DASHBOARD", font=fs(34), pad_x=50, pad_y=22)
    tc(d, (SIZE // 2, 232), "How Much Are You ACTUALLY Making?", fserif(50), WHITE)
    tc(d, (SIZE // 2, 300), "10 live KPIs + revenue-vs-profit, fees, product & ads charts — all automatic",
       fs(24, bold=False), (226, 214, 190))
    app_window(img, (70, 400, SIZE - 70, SIZE - 70), 0, content_dashboard)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_orders(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=320)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 112, "THE PROFIT ENGINE", font=fs(36), pad_x=52, pad_y=22)
    tc(d, (SIZE // 2, 230), "Every Order, Real Profit — Automatically", fserif(46), WHITE)
    tc(d, (SIZE // 2, 292), "Enter the sale. Etsy fees & true net profit compute instantly. Refunds flag red.",
       fs(23, bold=False), (226, 214, 190))
    app_window(img, (80, 360, SIZE - 80, SIZE - 70), 1, content_orders)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_fees_products(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=300)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 110, "FEES & PRODUCTS", font=fs(36), pad_x=52, pad_y=22)
    tc(d, (SIZE // 2, 224), "Where Your Money Goes & What To Scale", fserif(44), WHITE)
    app_window(img, (60, 330, SIZE - 60, 1150), 3, content_fees)
    app_window(img, (60, 1180, SIZE - 60, SIZE - 60), 6, content_library)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_mobile(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=400)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 130, "WORKS EVERYWHERE", font=fs(38), pad_x=54, pad_y=22)
    tc(d, (SIZE // 2, 250), "Excel · Google Sheets · Mobile", fserif(56), WHITE)
    tc(d, (SIZE // 2, 320), "Check your real profit from your phone between orders — no accounting degree required",
       fs(23, bold=False), (226, 214, 190))
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
    od.text(((sx0 + sx1) // 2, sy0 + 56), "Etsy Profit", font=fserif(36), fill=GOLD_HI, anchor="mm")
    y = sy0 + 150
    cards = [("NET PROFIT", "$530", PRIMARY), ("PROFIT MARGIN", "70%", PRIMARY),
             ("ETSY FEES", "$127", ACCENT), ("BEST SELLER", "Mega Bundle", PRIMARY)]
    for lab, val, col in cards:
        cb = (sx0 + 30, y, sx1 - 30, y + 135)
        od.rounded_rectangle(cb, radius=18, fill=WHITE, outline=GOLD_LT, width=2)
        od.rounded_rectangle((cb[0] + 20, y, cb[2] - 20, y + 5), radius=2, fill=GOLD_LT)
        od.text((cb[0] + 26, y + 34), lab, font=fs(22), fill=ACCENT, anchor="lt")
        vf = fit_font(od, val, sx1 - sx0 - 110, 46)
        od.text((cb[0] + 26, y + 94), val, font=vf, fill=col, anchor="lm")
        y += 152
    od.text((sx0 + 40, y + 16), "STRATEGY SAYS", font=fs(22), fill=ACCENT, anchor="lt")
    y += 52
    for lab, state in [("Scale Mega Bundle (82% margin)", True), ("Bundle IG + Stickers", True),
                       ("Fix Sticker Pack listing", False), ("Set aside $132 for tax", False)]:
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
        ("03_dashboard.png", render_dashboard),
        ("04_orders.png", render_orders),
        ("05_fees_products.png", render_fees_products),
        ("06_mobile.png", render_mobile),
    ]
    for name, fn in targets:
        fn(os.path.join(out_dir, name))
        print(f"  ✓ {name}")
    print(f"Wrote {len(targets)} images to {out_dir}")


if __name__ == "__main__":
    main()

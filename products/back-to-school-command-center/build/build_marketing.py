"""Generate the Etsy marketing image set for Back-to-School Command Center.

Outputs six 2000x2000 PNGs to ../marketing/:
  01_hero.png        - main listing thumbnail (Driver-Budget format)
  02_dashboard.png   - dashboard close-up
  03_supplies.png    - supplies tracker close-up
  04_budget.png      - budget breakdown with charts
  05_schedule.png    - weekly class schedule
  06_mobile.png      - mobile Google Sheets preview

Run: python3 build_marketing.py
"""
from __future__ import annotations

import math
import os
import random
from PIL import Image, ImageDraw, ImageFilter, ImageFont

# ---------------------------------------------------------------------------
# Brand tokens
# ---------------------------------------------------------------------------
PRIMARY = (27, 79, 72)
ACCENT = (147, 115, 86)
SURFACE = (229, 211, 186)
HIGHLIGHT = (117, 230, 193)
BG = (252, 250, 246)
WHITE = (255, 255, 255)
TEXT = (51, 51, 51)
TEXT_MUTED = (120, 120, 120)
DANGER = (201, 76, 76)
SOFT_BG = (250, 247, 241)
DOT = (220, 215, 205)
WARN_BG = (251, 240, 226)
MINT_BG = (227, 248, 239)
RED_BG = (251, 230, 230)

FONT_BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
FONT_REG = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"

SIZE = 2000


def f(size: int, bold: bool = True) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(FONT_BOLD if bold else FONT_REG, size)


# ---------------------------------------------------------------------------
# Primitives
# ---------------------------------------------------------------------------
def dotted_bg(canvas: Image.Image, dot_color=DOT, spacing: int = 44, radius: int = 4) -> None:
    d = ImageDraw.Draw(canvas)
    for y in range(spacing // 2, canvas.height, spacing):
        for x in range(spacing // 2, canvas.width, spacing):
            d.ellipse((x - radius, y - radius, x + radius, y + radius), fill=dot_color)


def drop_shadow(canvas: Image.Image, box, radius: int, blur: int = 24, alpha: int = 70):
    layer = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    ImageDraw.Draw(layer).rounded_rectangle(box, radius=radius, fill=(27, 79, 72, alpha))
    canvas.alpha_composite(layer.filter(ImageFilter.GaussianBlur(blur)), (0, 18))


def text_size(draw, text, font):
    box = draw.textbbox((0, 0), text, font=font)
    return box[2] - box[0], box[3] - box[1]


def text_center(draw, xy, text, font, fill, anchor="mm"):
    draw.text(xy, text, font=font, fill=fill, anchor=anchor)


def pill(canvas, cx, cy, text, font, fg=WHITE, bg=PRIMARY,
         pad_x=60, pad_y=26, star=False):
    overlay = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(overlay)
    label = f"★ {text}" if star else text
    tw, th = text_size(d, label, font)
    w, h = tw + pad_x * 2, th + pad_y * 2
    box = (cx - w // 2, cy - h // 2, cx + w // 2, cy + h // 2)
    d.rounded_rectangle(box, radius=h // 2, fill=bg)
    d.text((cx, cy), label, font=font, fill=fg, anchor="mm")
    canvas.alpha_composite(overlay)


def kpi_callout(canvas, cx, cy, label, value, width=720, height=280,
                value_color=None):
    box = (cx - width // 2, cy - height // 2, cx + width // 2, cy + height // 2)
    drop_shadow(canvas, box, radius=24, blur=30, alpha=85)
    overlay = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(overlay)
    d.rounded_rectangle(box, radius=24, fill=WHITE, outline=(220, 215, 205), width=2)
    d.text((cx, cy - 70), label, font=f(46), fill=TEXT, anchor="mm")
    d.text((cx, cy + 50), value, font=f(108), fill=value_color or PRIMARY, anchor="mm")
    canvas.alpha_composite(overlay)


# ---------------------------------------------------------------------------
# Faux dashboard renderer for the laptop screen
# ---------------------------------------------------------------------------
def draw_dashboard_in_screen(canvas, screen):
    sx0, sy0, sx1, sy1 = screen
    sw = sx1 - sx0
    sh = sy1 - sy0
    overlay = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(overlay)

    # Top header
    head_h = 86
    d.rounded_rectangle((sx0, sy0, sx1, sy0 + head_h), radius=8, fill=PRIMARY)
    d.rectangle((sx0, sy0 + 30, sx1, sy0 + head_h), fill=PRIMARY)
    d.text((sx0 + 28, sy0 + head_h // 2), "BACK-TO-SCHOOL COMMAND CENTER",
           font=f(26), fill=WHITE, anchor="lm")
    d.text((sx1 - 28, sy0 + head_h // 2), "2026 – 2027 SCHOOL YEAR",
           font=f(22), fill=WHITE, anchor="rm")

    # KPI grid - 4 cards across, 2 rows
    gx = sx0 + 24
    gy = sy0 + head_h + 24
    gap = 18
    kw = (sw - 24 * 2 - gap * 3) // 4
    kh = 130

    kpis = [
        ("TOTAL BUDGET",      "$1,200",  PRIMARY),
        ("TOTAL SPENT",       "$558",    ACCENT),
        ("REMAINING",         "$642",    PRIMARY),
        ("DAYS UNTIL SCHOOL", "47",      DANGER),
        ("ITEMS BOUGHT",      "9",       PRIMARY),
        ("ITEMS LEFT",        "13",      ACCENT),
        ("SUPPLY %",          "41%",     PRIMARY),
        ("ASSIGNMENTS",       "30%",     PRIMARY),
    ]
    for i, (lab, val, col) in enumerate(kpis):
        row, col_i = divmod(i, 4)
        x0 = gx + col_i * (kw + gap)
        y0 = gy + row * (kh + gap)
        box = (x0, y0, x0 + kw, y0 + kh)
        d.rounded_rectangle(box, radius=10, fill=WHITE, outline=(220, 215, 205), width=2)
        d.text((x0 + 14, y0 + 16), lab, font=f(18), fill=ACCENT, anchor="lt")
        d.text((x0 + 14, y0 + 76), val, font=f(46), fill=col, anchor="lm")

    # Quick nav chip row
    nav_y = gy + 2 * (kh + gap) + 10
    nav_h = 36
    nav = ["Shopping", "Budget", "Calendar", "Students",
           "Assignments", "Meals", "Emergency"]
    cw = (sw - 24 * 2 - 6 * (len(nav) - 1)) / len(nav)
    for i, name in enumerate(nav):
        x0 = gx + i * (cw + 6)
        d.rounded_rectangle((x0, nav_y, x0 + cw, nav_y + nav_h),
                            radius=18, fill=PRIMARY)
        d.text((x0 + cw / 2, nav_y + nav_h / 2), name,
               font=f(16), fill=WHITE, anchor="mm")

    # Two-chart row (donut + column)
    ch_y0 = nav_y + nav_h + 22
    ch_h = sy1 - ch_y0 - 20
    chart_w = (sw - 24 * 2 - gap) // 2

    # Donut: Spending by Category
    dbox = (gx, ch_y0, gx + chart_w, ch_y0 + ch_h)
    d.rounded_rectangle(dbox, radius=10, fill=WHITE, outline=(220, 215, 205), width=2)
    d.text((dbox[0] + 16, dbox[1] + 14), "SPENDING BY CATEGORY",
           font=f(20), fill=ACCENT, anchor="lt")
    dcx = dbox[0] + 110
    dcy = (dbox[1] + dbox[3]) // 2 + 12
    rr = 80
    segs = [
        (35, PRIMARY),
        (22, ACCENT),
        (15, HIGHLIGHT),
        (12, SURFACE),
        (10, (180, 90, 90)),
        (6,  TEXT_MUTED),
    ]
    angle = -90
    for pct, col in segs:
        sw_ = pct * 3.6
        d.pieslice((dcx - rr, dcy - rr, dcx + rr, dcy + rr),
                   angle, angle + sw_, fill=col)
        angle += sw_
    d.ellipse((dcx - 36, dcy - 36, dcx + 36, dcy + 36), fill=WHITE)
    d.text((dcx, dcy - 6), "$558", font=f(20), fill=PRIMARY, anchor="mm")
    d.text((dcx, dcy + 14), "SPENT", font=f(11), fill=TEXT_MUTED, anchor="mm")
    # Legend
    legend = [("Supplies", PRIMARY), ("Clothing", ACCENT),
              ("Electronics", HIGHLIGHT), ("Shoes", SURFACE),
              ("Fees", (180, 90, 90)), ("Other", TEXT_MUTED)]
    lx = dcx + rr + 30
    ly = dcy - rr + 6
    for i, (label, col) in enumerate(legend):
        yy = ly + i * 24
        d.rounded_rectangle((lx, yy, lx + 16, yy + 16), radius=3, fill=col)
        d.text((lx + 26, yy + 8), label, font=f(16), fill=TEXT, anchor="lm")

    # Column chart: Budget vs Actual
    bbox = (gx + chart_w + gap, ch_y0, gx + 2 * chart_w + gap, ch_y0 + ch_h)
    d.rounded_rectangle(bbox, radius=10, fill=WHITE, outline=(220, 215, 205), width=2)
    d.text((bbox[0] + 16, bbox[1] + 14), "BUDGET vs ACTUAL",
           font=f(20), fill=ACCENT, anchor="lt")
    cats = ["Supplies", "Cloth", "Shoes", "Elec", "Sports", "Lunch", "Fees"]
    plan = [200, 260, 120, 180, 90, 60, 180]
    actual = [115, 145, 40, 95, 0, 35, 0]
    bx0, by0 = bbox[0] + 30, bbox[1] + 60
    bx1, by1 = bbox[2] - 20, bbox[3] - 50
    max_v = max(plan)
    bw = (bx1 - bx0) / (len(cats) * 2.3)
    for i, (p, a) in enumerate(zip(plan, actual)):
        x = bx0 + i * (bw * 2.3) + 12
        h_p = (p / max_v) * (by1 - by0)
        h_a = (a / max_v) * (by1 - by0)
        d.rounded_rectangle((x, by1 - h_p, x + bw, by1), radius=4, fill=SURFACE)
        d.rounded_rectangle((x + bw + 4, by1 - h_a, x + 2 * bw + 4, by1),
                            radius=4, fill=PRIMARY)
        d.text((x + bw + 2, by1 + 18), cats[i],
               font=f(14), fill=TEXT_MUTED, anchor="mt")

    canvas.alpha_composite(overlay)


def draw_laptop(canvas, cx, cy, w=1500, h=900):
    overlay = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(overlay)
    # Shadow
    sbox = (cx - w // 2, cy - h // 2 + 30, cx + w // 2, cy + h // 2 + 80)
    shadow = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    ImageDraw.Draw(shadow).rounded_rectangle(sbox, radius=40, fill=(27, 79, 72, 110))
    canvas.alpha_composite(shadow.filter(ImageFilter.GaussianBlur(50)))
    # Body
    body = (cx - w // 2, cy - h // 2, cx + w // 2, cy + h // 2)
    d.rounded_rectangle(body, radius=28, fill=(40, 40, 45), outline=(20, 20, 25), width=4)
    bezel = 22
    screen = (body[0] + bezel, body[1] + bezel, body[2] - bezel, body[3] - bezel)
    d.rounded_rectangle(screen, radius=8, fill=BG)
    # Base
    d.rounded_rectangle((cx - w // 2 - 120, cy + h // 2,
                         cx + w // 2 + 120, cy + h // 2 + 24),
                        radius=12, fill=(60, 60, 65))
    d.rounded_rectangle((cx - 80, cy + h // 2 + 4, cx + 80, cy + h // 2 + 14),
                        radius=6, fill=(35, 35, 40))
    canvas.alpha_composite(overlay)
    draw_dashboard_in_screen(canvas, screen)


# ===========================================================================
# Image 1: HERO
# ===========================================================================
def render_hero(out: str) -> None:
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)

    pill(img, SIZE // 2, 230, "ULTIMATE 12-DASHBOARD SYSTEM",
         font=f(46), pad_x=70, pad_y=28)

    d = ImageDraw.Draw(img)
    text_center(d, (SIZE // 2, 430), "BACK-TO-SCHOOL", f(170), PRIMARY)
    text_center(d, (SIZE // 2, 590), "COMMAND CENTER", f(160), PRIMARY)

    # Subtitle
    sub_font = f(50)
    parts = ["PLAN", "BUDGET", "ORGANIZE"]
    full = "  •  ".join(parts)
    tw, _ = text_size(d, full, sub_font)
    text_center(d, (SIZE // 2, 720), full, sub_font, TEXT)
    for sx in (SIZE // 2 - tw // 2 - 70, SIZE // 2 + tw // 2 + 70):
        d.polygon([(sx, 700), (sx + 20, 720), (sx, 740), (sx - 20, 720)], fill=PRIMARY)

    # Laptop with dashboard
    draw_laptop(img, SIZE // 2, 1240, w=1640, h=940)

    # Two floating KPI callouts
    kpi_callout(img, 360, 1010, "DAYS TO SCHOOL", "47",
                width=680, height=260, value_color=DANGER)
    kpi_callout(img, SIZE - 360, 1010, "REMAINING", "$642",
                width=680, height=260)

    # Bottom CTA
    pill(img, SIZE // 2, SIZE - 170,
         "GET THE 12-DASHBOARD SYSTEM",
         font=f(58), pad_x=80, pad_y=34, star=True)

    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
# Image 2: DASHBOARD CLOSEUP
# ===========================================================================
def render_dashboard_closeup(out: str) -> None:
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)

    pill(img, SIZE // 2, 180, "LIVE DASHBOARD PREVIEW",
         font=f(44), pad_x=60, pad_y=24)

    d = ImageDraw.Draw(img)
    text_center(d, (SIZE // 2, 320), "8 KPIs · 2 CHARTS · ONE SCREEN", f(74), PRIMARY)
    text_center(d, (SIZE // 2, 410),
                "Track budget, supplies, assignments & countdown — all live",
                f(34, bold=False), TEXT_MUTED)

    panel = (140, 500, SIZE - 140, SIZE - 200)
    drop_shadow(img, panel, radius=24, blur=40, alpha=110)
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    od.rounded_rectangle(panel, radius=24, fill=BG, outline=(220, 215, 205), width=3)
    img.alpha_composite(overlay)
    inner = (panel[0] + 24, panel[1] + 24, panel[2] - 24, panel[3] - 24)
    draw_dashboard_in_screen(img, inner)

    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
# Image 3: SUPPLIES TRACKER
# ===========================================================================
def render_supplies(out: str) -> None:
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)
    d = ImageDraw.Draw(img)

    pill(img, SIZE // 2, 180, "SCHOOL SUPPLY TRACKER",
         font=f(44), pad_x=60, pad_y=24)
    text_center(d, (SIZE // 2, 320), "EVERY ITEM. EVERY STORE.", f(76), PRIMARY)
    text_center(d, (SIZE // 2, 410),
                "Validated dropdowns, priority bars, auto-tally — never miss a thing",
                f(32, bold=False), TEXT_MUTED)

    panel = (120, 510, SIZE - 120, SIZE - 180)
    drop_shadow(img, panel, radius=24, blur=40, alpha=110)
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    od.rounded_rectangle(panel, radius=24, fill=WHITE, outline=(220, 215, 205), width=3)
    img.alpha_composite(overlay)

    headers = ["CATEGORY", "ITEM", "REQ", "BOUGHT", "LEFT", "STORE", "PRICE", "PURCHASED", "PRIORITY"]
    rows = [
        ("Supplies",    "Backpack",            1,  1, 0, "Target",  "$34.99", "Yes", "High",   MINT_BG),
        ("Supplies",    "Pencils #2 (12pk)",   4,  3, 1, "Walmart", "$3.49",  "No",  "High",   WARN_BG),
        ("Supplies",    "Glue sticks",        12,  6, 6, "Staples", "$0.99",  "No",  "Medium", BG),
        ("Supplies",    "Notebook (wide-rule)",8,  8, 0, "Target",  "$1.29",  "Yes", "High",   MINT_BG),
        ("Supplies",    "Folders",             6,  3, 3, "Target",  "$0.79",  "No",  "Medium", BG),
        ("Electronics", "USB-C charger",       2,  1, 1, "Amazon", "$19.99", "No",  "High",   WARN_BG),
        ("Electronics", "Headphones",          2,  2, 0, "Target", "$12.99", "Yes", "Medium", MINT_BG),
        ("Clothing",    "School polo",         5,  2, 3, "Target", "$14.99", "No",  "High",   WARN_BG),
        ("Shoes",       "Sneakers",            2,  1, 1, "Local",  "$39.99", "No",  "High",   WARN_BG),
        ("Fees",        "Registration fee",    1,  0, 1, "Other", "$125.00", "No",  "High",   RED_BG),
    ]
    inner = (panel[0] + 24, panel[1] + 24, panel[2] - 24, panel[3] - 24)
    iw = inner[2] - inner[0]
    ih = inner[3] - inner[1]
    cols = len(headers)
    col_w = iw / cols
    hdr_h = 78

    overlay2 = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay2)
    od.rounded_rectangle((inner[0], inner[1], inner[2], inner[1] + hdr_h),
                         radius=14, fill=PRIMARY)
    od.rectangle((inner[0], inner[1] + 30, inner[2], inner[1] + hdr_h), fill=PRIMARY)
    for i, h in enumerate(headers):
        cx = inner[0] + col_w * (i + 0.5)
        od.text((cx, inner[1] + hdr_h // 2), h, font=f(22), fill=WHITE, anchor="mm")

    row_h = (ih - hdr_h - 12) / len(rows)
    for ri, row in enumerate(rows):
        y0 = inner[1] + hdr_h + ri * row_h
        bg_fill = row[-1]
        if bg_fill == BG and ri % 2 == 1:
            bg_fill = (244, 236, 222)
        if bg_fill != BG:
            od.rectangle((inner[0] + 4, y0, inner[2] - 4, y0 + row_h), fill=bg_fill)
        # Cells
        for ci, val in enumerate(row[:-1]):
            cx = inner[0] + col_w * (ci + 0.5)
            color = PRIMARY if ci == 0 else TEXT
            font_use = f(24) if ci != 0 else f(24)
            # Status pill for Purchased + Priority columns
            if ci == 7:  # Purchased
                pill_color = HIGHLIGHT if val == "Yes" else SURFACE
                tx_color = PRIMARY
                pw, ph = 80, 36
                od.rounded_rectangle((cx - pw // 2, y0 + row_h / 2 - ph // 2,
                                      cx + pw // 2, y0 + row_h / 2 + ph // 2),
                                     radius=18, fill=pill_color)
                od.text((cx, y0 + row_h / 2), val, font=f(20), fill=tx_color, anchor="mm")
            elif ci == 8:  # Priority
                pri_color = DANGER if val == "High" else ACCENT if val == "Medium" else TEXT_MUTED
                pw, ph = 100, 36
                od.rounded_rectangle((cx - pw // 2, y0 + row_h / 2 - ph // 2,
                                      cx + pw // 2, y0 + row_h / 2 + ph // 2),
                                     radius=18, fill=pri_color)
                od.text((cx, y0 + row_h / 2), val, font=f(20), fill=WHITE, anchor="mm")
            else:
                od.text((cx, y0 + row_h / 2), str(val), font=font_use, fill=color, anchor="mm")

    img.alpha_composite(overlay2)
    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
# Image 4: BUDGET CHARTS
# ===========================================================================
def render_budget(out: str) -> None:
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)
    d = ImageDraw.Draw(img)

    pill(img, SIZE // 2, 180, "BACK-TO-SCHOOL BUDGET",
         font=f(44), pad_x=60, pad_y=24)
    text_center(d, (SIZE // 2, 320), "PLAN. TRACK. STAY ON BUDGET.", f(70), PRIMARY)
    text_center(d, (SIZE // 2, 410),
                "Auto-pulled actuals from your supply list — variance & % spent live",
                f(32, bold=False), TEXT_MUTED)

    # KPI strip top
    kpis = [("PLANNED", "$1,200"), ("SPENT", "$558"),
            ("REMAINING", "$642"), ("% SPENT", "47%")]
    kw, kh = 380, 180
    gap = 30
    total_w = 4 * kw + 3 * gap
    start_x = (SIZE - total_w) // 2
    ky = 510
    for i, (lab, val) in enumerate(kpis):
        x = start_x + i * (kw + gap)
        box = (x, ky, x + kw, ky + kh)
        drop_shadow(img, box, radius=20, blur=24, alpha=80)
        overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
        od = ImageDraw.Draw(overlay)
        od.rounded_rectangle(box, radius=20, fill=WHITE, outline=(220, 215, 205), width=2)
        od.text((x + kw // 2, ky + 50), lab, font=f(28), fill=ACCENT, anchor="mm")
        od.text((x + kw // 2, ky + 120), val, font=f(70), fill=PRIMARY, anchor="mm")
        img.alpha_composite(overlay)

    # Donut card
    dbox = (140, 760, 950, SIZE - 200)
    drop_shadow(img, dbox, radius=24, blur=30, alpha=90)
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    od.rounded_rectangle(dbox, radius=24, fill=WHITE, outline=(220, 215, 205), width=2)
    od.text((dbox[0] + 30, dbox[1] + 30), "SPENDING BREAKDOWN", font=f(30), fill=ACCENT, anchor="lt")
    dcx = (dbox[0] + dbox[2]) // 2 - 100
    dcy = (dbox[1] + dbox[3]) // 2 + 30
    rr = 220
    segs = [
        (35, PRIMARY,  "Supplies 35%"),
        (22, ACCENT,   "Clothing 22%"),
        (15, HIGHLIGHT,"Electronics 15%"),
        (12, SURFACE,  "Shoes 12%"),
        (10, (180, 90, 90), "Fees 10%"),
        (6,  TEXT_MUTED, "Other 6%"),
    ]
    angle = -90
    for pct, col, _ in segs:
        sweep = pct * 3.6
        od.pieslice((dcx - rr, dcy - rr, dcx + rr, dcy + rr),
                    angle, angle + sweep, fill=col)
        angle += sweep
    od.ellipse((dcx - 90, dcy - 90, dcx + 90, dcy + 90), fill=WHITE)
    od.text((dcx, dcy - 12), "$558", font=f(44), fill=PRIMARY, anchor="mm")
    od.text((dcx, dcy + 30), "SPENT", font=f(20), fill=TEXT_MUTED, anchor="mm")
    # Legend
    lx = dcx + rr + 40
    for i, (pct, col, label) in enumerate(segs):
        yy = dcy - rr + 20 + i * 50
        od.rounded_rectangle((lx, yy, lx + 28, yy + 28), radius=5, fill=col)
        od.text((lx + 44, yy + 14), label, font=f(24), fill=TEXT, anchor="lm")
    img.alpha_composite(overlay)

    # Bar chart card
    bbox = (1010, 760, SIZE - 140, SIZE - 200)
    drop_shadow(img, bbox, radius=24, blur=30, alpha=90)
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    od.rounded_rectangle(bbox, radius=24, fill=WHITE, outline=(220, 215, 205), width=2)
    od.text((bbox[0] + 30, bbox[1] + 30), "BUDGET vs ACTUAL", font=f(30), fill=ACCENT, anchor="lt")
    cats = ["Supplies", "Clothing", "Shoes", "Electronics", "Sports", "Lunch", "Fees", "Misc"]
    plan = [200, 260, 120, 180, 90, 60, 180, 50]
    actual = [115, 145, 40, 95, 0, 35, 0, 12]
    bx0, by0 = bbox[0] + 50, bbox[1] + 110
    bx1, by1 = bbox[2] - 40, bbox[3] - 90
    max_v = max(plan)
    bw = (bx1 - bx0) / (len(cats) * 2.4)
    for i, (p, a) in enumerate(zip(plan, actual)):
        x = bx0 + i * (bw * 2.4) + 20
        h_p = (p / max_v) * (by1 - by0)
        h_a = (a / max_v) * (by1 - by0)
        od.rounded_rectangle((x, by1 - h_p, x + bw, by1), radius=6, fill=SURFACE)
        od.rounded_rectangle((x + bw + 8, by1 - h_a, x + 2 * bw + 8, by1),
                             radius=6, fill=PRIMARY)
        od.text((x + bw + 4, by1 + 22), cats[i],
                font=f(18), fill=TEXT_MUTED, anchor="mt")
    # Legend
    od.rounded_rectangle((bbox[0] + 30, bbox[3] - 50, bbox[0] + 50, bbox[3] - 30),
                         radius=4, fill=SURFACE)
    od.text((bbox[0] + 60, bbox[3] - 40), "Planned", font=f(20), fill=TEXT, anchor="lm")
    od.rounded_rectangle((bbox[0] + 180, bbox[3] - 50, bbox[0] + 200, bbox[3] - 30),
                         radius=4, fill=PRIMARY)
    od.text((bbox[0] + 210, bbox[3] - 40), "Actual", font=f(20), fill=TEXT, anchor="lm")
    img.alpha_composite(overlay)

    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
# Image 5: WEEKLY SCHEDULE
# ===========================================================================
def render_schedule(out: str) -> None:
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)
    d = ImageDraw.Draw(img)

    pill(img, SIZE // 2, 180, "CLASS SCHEDULE",
         font=f(44), pad_x=60, pad_y=24)
    text_center(d, (SIZE // 2, 320), "COLOR-CODED WEEKLY VIEW", f(76), PRIMARY)
    text_center(d, (SIZE // 2, 410),
                "See every class, every period, every day at a glance",
                f(32, bold=False), TEXT_MUTED)

    # Schedule grid
    panel = (120, 510, SIZE - 120, SIZE - 180)
    drop_shadow(img, panel, radius=24, blur=40, alpha=110)
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    od.rounded_rectangle(panel, radius=24, fill=WHITE, outline=(220, 215, 205), width=3)
    img.alpha_composite(overlay)
    inner = (panel[0] + 28, panel[1] + 28, panel[2] - 28, panel[3] - 28)

    days = ["TIME", "MON", "TUE", "WED", "THU", "FRI"]
    times = ["8:00", "9:00", "10:00", "11:00", "12:00", "1:00", "2:00"]
    subjects_grid = [
        ["Math",     "Math",     "Math",     "Math",     "Math"],
        ["English",  "Science",  "English",  "Science",  "English"],
        ["Science",  "English",  "Science",  "English",  "Science"],
        ["History",  "Art",      "History",  "Tech",     "Music"],
        ["LUNCH",    "LUNCH",    "LUNCH",    "LUNCH",    "LUNCH"],
        ["PE",       "PE",       "PE",       "PE",       "PE"],
        ["Spanish",  "Spanish",  "Spanish",  "Spanish",  "Spanish"],
    ]
    rooms = [
        ["R104",  "R104",  "R104",  "R104",  "R104"],
        ["R208",  "R301",  "R208",  "R301",  "R208"],
        ["R301",  "R208",  "R301",  "R208",  "R301"],
        ["R112",  "R150",  "R112",  "Lab2",  "R220"],
        ["—",     "—",     "—",     "—",     "—"   ],
        ["Gym",   "Gym",   "Gym",   "Gym",   "Gym" ],
        ["R141",  "R141",  "R141",  "R141",  "R141"],
    ]
    subject_colors = {
        "Math":     (217, 232, 230),
        "English":  (237, 224, 206),
        "Science":  (220, 245, 236),
        "History":  (240, 228, 210),
        "Art":      (245, 224, 214),
        "Tech":     (221, 229, 242),
        "Music":    (234, 217, 240),
        "PE":       (245, 226, 226),
        "Spanish":  (224, 235, 225),
        "LUNCH":    SURFACE,
    }

    iw = inner[2] - inner[0]
    ih = inner[3] - inner[1]
    col_w = iw / len(days)
    row_h = ih / (len(times) + 1)

    overlay2 = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay2)
    # Header row
    od.rounded_rectangle((inner[0], inner[1], inner[2], inner[1] + row_h),
                         radius=12, fill=PRIMARY)
    od.rectangle((inner[0], inner[1] + 20, inner[2], inner[1] + row_h), fill=PRIMARY)
    for i, day in enumerate(days):
        cx = inner[0] + col_w * (i + 0.5)
        od.text((cx, inner[1] + row_h / 2), day, font=f(30), fill=WHITE, anchor="mm")
    # Body
    for ri in range(len(times)):
        y0 = inner[1] + row_h * (ri + 1)
        # Time column
        od.rectangle((inner[0], y0, inner[0] + col_w, y0 + row_h), fill=PRIMARY)
        od.text((inner[0] + col_w / 2, y0 + row_h / 2),
                times[ri], font=f(26), fill=WHITE, anchor="mm")
        for ci in range(5):
            x0 = inner[0] + col_w * (ci + 1)
            subj = subjects_grid[ri][ci]
            room = rooms[ri][ci]
            color = subject_colors.get(subj, SOFT_BG)
            od.rectangle((x0 + 2, y0 + 2, x0 + col_w - 2, y0 + row_h - 2),
                         fill=color, outline=(220, 215, 205))
            od.text((x0 + col_w / 2, y0 + row_h / 2 - 14),
                    subj, font=f(26), fill=PRIMARY, anchor="mm")
            if room != "—":
                od.text((x0 + col_w / 2, y0 + row_h / 2 + 18),
                        room, font=f(20), fill=TEXT_MUTED, anchor="mm")
    img.alpha_composite(overlay2)

    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
# Image 6: MOBILE PREVIEW
# ===========================================================================
def render_mobile(out: str) -> None:
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)
    d = ImageDraw.Draw(img)

    pill(img, SIZE // 2, 180, "WORKS EVERYWHERE",
         font=f(44), pad_x=60, pad_y=24)
    text_center(d, (SIZE // 2, 330), "EXCEL  ·  GOOGLE SHEETS", f(78), PRIMARY)
    text_center(d, (SIZE // 2, 420),
                "Update on phone after pickup — syncs to the family dashboard",
                f(32, bold=False), TEXT_MUTED)

    # Phone frame
    px, py = SIZE // 2, 1280
    pw, ph = 640, 1240
    phone_box = (px - pw // 2, py - ph // 2, px + pw // 2, py + ph // 2)
    drop_shadow(img, phone_box, radius=60, blur=50, alpha=120)
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    od.rounded_rectangle(phone_box, radius=60, fill=(25, 25, 30))
    bezel = 22
    screen = (phone_box[0] + bezel, phone_box[1] + bezel + 30,
              phone_box[2] - bezel, phone_box[3] - bezel - 30)
    od.rounded_rectangle(screen, radius=42, fill=BG)
    od.rounded_rectangle((px - 90, phone_box[1] + 16, px + 90, phone_box[1] + 48),
                         radius=18, fill=(15, 15, 18))

    sx0, sy0, sx1, sy1 = screen
    # Top bar
    od.rounded_rectangle((sx0, sy0, sx1, sy0 + 90), radius=42, fill=PRIMARY)
    od.rectangle((sx0, sy0 + 50, sx1, sy0 + 90), fill=PRIMARY)
    od.text(((sx0 + sx1) // 2, sy0 + 50), "Command Center",
            font=f(32), fill=WHITE, anchor="mm")

    # KPI cards (stacked)
    y = sy0 + 130
    kpis_mobile = [
        ("DAYS UNTIL SCHOOL", "47",      DANGER),
        ("REMAINING BUDGET",  "$642",    PRIMARY),
        ("SUPPLY COMPLETION", "41%",     PRIMARY),
        ("ASSIGNMENTS DONE",  "30%",     PRIMARY),
    ]
    for lab, val, col in kpis_mobile:
        cbox = (sx0 + 30, y, sx1 - 30, y + 140)
        od.rounded_rectangle(cbox, radius=18, fill=WHITE, outline=(220, 215, 205), width=2)
        od.text((cbox[0] + 26, y + 30), lab, font=f(22), fill=ACCENT, anchor="lt")
        od.text((cbox[0] + 26, y + 96), val, font=f(48), fill=col, anchor="lm")
        y += 160

    # Mini upcoming events list
    od.text((sx0 + 40, y + 20), "UPCOMING", font=f(22), fill=ACCENT, anchor="lt")
    y += 50
    events = [
        ("Sep 02", "First day of school", HIGHLIGHT),
        ("Sep 14", "Picture day",         SURFACE),
        ("Sep 21", "Back-to-school night",SURFACE),
        ("Oct 09", "Teacher work day",    SURFACE),
    ]
    for date_str, evt, dot_col in events:
        od.ellipse((sx0 + 40, y + 14, sx0 + 64, y + 38), fill=dot_col)
        od.text((sx0 + 80, y + 14), date_str, font=f(20), fill=PRIMARY, anchor="lt")
        od.text((sx0 + 80, y + 40), evt, font=f(22), fill=TEXT, anchor="lt")
        y += 76

    img.alpha_composite(overlay)
    img.convert("RGB").save(out, "PNG", optimize=True)


# ---------------------------------------------------------------------------
# Build all
# ---------------------------------------------------------------------------
def main() -> None:
    out_dir = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "marketing"
    )
    os.makedirs(out_dir, exist_ok=True)
    targets = [
        ("01_hero.png",      render_hero),
        ("02_dashboard.png", render_dashboard_closeup),
        ("03_supplies.png",  render_supplies),
        ("04_budget.png",    render_budget),
        ("05_schedule.png",  render_schedule),
        ("06_mobile.png",    render_mobile),
    ]
    for name, fn in targets:
        path = os.path.join(out_dir, name)
        fn(path)
        print(f"  ✓ {name}")
    print(f"Wrote {len(targets)} images to {out_dir}")


if __name__ == "__main__":
    main()

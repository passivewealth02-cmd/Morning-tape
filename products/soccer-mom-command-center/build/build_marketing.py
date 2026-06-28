"""Generate the Etsy marketing image set for Soccer Mom Command Center™.

Outputs six 2000x2000 PNGs to ../marketing/:
  01_hero.png        - main listing thumbnail (Driver-Budget format)
  02_dashboard.png   - dashboard close-up
  03_schedule.png    - season schedule with Days-Out
  04_budget.png      - budget breakdown with charts
  05_tournament.png  - tournament planner + packing checklist
  06_mobile.png      - mobile Google Sheets preview

Run: python3 build_marketing.py
"""
from __future__ import annotations

import math
import os
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


def drop_shadow(canvas, box, radius, blur=24, alpha=70):
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
# Faux SMCC dashboard inside laptop / panel
# ---------------------------------------------------------------------------
def draw_dashboard_in_screen(canvas, screen):
    sx0, sy0, sx1, sy1 = screen
    sw = sx1 - sx0
    sh = sy1 - sy0
    overlay = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(overlay)

    # Header
    head_h = 86
    d.rounded_rectangle((sx0, sy0, sx1, sy0 + head_h), radius=8, fill=PRIMARY)
    d.rectangle((sx0, sy0 + 30, sx1, sy0 + head_h), fill=PRIMARY)
    d.text((sx0 + 28, sy0 + head_h // 2), "SOCCER MOM COMMAND CENTER",
           font=f(26), fill=WHITE, anchor="lm")
    d.text((sx1 - 28, sy0 + head_h // 2), "FALL 2026 · U10 TRAVEL",
           font=f(22), fill=WHITE, anchor="rm")

    # KPI grid 4x2
    gx = sx0 + 24
    gy = sy0 + head_h + 24
    gap = 18
    kw = (sw - 24 * 2 - gap * 3) // 4
    kh = 130

    kpis = [
        ("GAMES THIS MONTH",   "5",       PRIMARY),
        ("PRACTICES THIS WK",  "3",       PRIMARY),
        ("MONTHLY BUDGET",     "$650",    PRIMARY),
        ("REMAINING",          "$214",    HIGHLIGHT if True else PRIMARY),
        ("UPCOMING TOURNEYS",  "2",       ACCENT),
        ("EQUIPMENT NEEDED",   "1",       ACCENT),
        ("FAMILY CONFLICTS",   "0",       PRIMARY),
        ("ATTENDANCE %",       "92%",     PRIMARY),
    ]
    for i, (lab, val, col) in enumerate(kpis):
        row, col_i = divmod(i, 4)
        x0 = gx + col_i * (kw + gap)
        y0 = gy + row * (kh + gap)
        box = (x0, y0, x0 + kw, y0 + kh)
        d.rounded_rectangle(box, radius=10, fill=WHITE, outline=(220, 215, 205), width=2)
        d.text((x0 + 14, y0 + 16), lab, font=f(16), fill=ACCENT, anchor="lt")
        if col == HIGHLIGHT:
            d.text((x0 + 14, y0 + 76), val, font=f(44), fill=PRIMARY, anchor="lm")
        else:
            d.text((x0 + 14, y0 + 76), val, font=f(46), fill=col, anchor="lm")

    # Quick nav chip row
    nav_y = gy + 2 * (kh + gap) + 10
    nav_h = 36
    nav = ["Schedule", "Budget", "Equipment", "Roster",
           "Travel", "Meals", "Carpool", "Comms"]
    cw = (sw - 24 * 2 - 6 * (len(nav) - 1)) / len(nav)
    for i, name in enumerate(nav):
        x0 = gx + i * (cw + 6)
        d.rounded_rectangle((x0, nav_y, x0 + cw, nav_y + nav_h),
                            radius=18, fill=PRIMARY)
        d.text((x0 + cw / 2, nav_y + nav_h / 2), name,
               font=f(16), fill=WHITE, anchor="mm")

    # Two-chart row (donut + bar)
    ch_y0 = nav_y + nav_h + 22
    ch_h = sy1 - ch_y0 - 20
    chart_w = (sw - 24 * 2 - gap) // 2

    # Donut
    dbox = (gx, ch_y0, gx + chart_w, ch_y0 + ch_h)
    d.rounded_rectangle(dbox, radius=10, fill=WHITE, outline=(220, 215, 205), width=2)
    d.text((dbox[0] + 16, dbox[1] + 14), "SPENDING BY CATEGORY",
           font=f(20), fill=ACCENT, anchor="lt")
    dcx = dbox[0] + 110
    dcy = (dbox[1] + dbox[3]) // 2 + 12
    rr = 80
    segs = [
        (22, PRIMARY),       # Tournaments
        (16, ACCENT),        # Hotels
        (14, HIGHLIGHT),     # Registration
        (12, SURFACE),       # Uniforms
        (10, (180, 90, 90)), # Coaching
        (26, TEXT_MUTED),    # Other
    ]
    angle = -90
    for pct, col in segs:
        sw_ = pct * 3.6
        d.pieslice((dcx - rr, dcy - rr, dcx + rr, dcy + rr),
                   angle, angle + sw_, fill=col)
        angle += sw_
    d.ellipse((dcx - 36, dcy - 36, dcx + 36, dcy + 36), fill=WHITE)
    d.text((dcx, dcy - 6), "$436", font=f(20), fill=PRIMARY, anchor="mm")
    d.text((dcx, dcy + 14), "SPENT", font=f(11), fill=TEXT_MUTED, anchor="mm")
    # Legend
    legend = [("Tournaments", PRIMARY), ("Hotels", ACCENT),
              ("Registration", HIGHLIGHT), ("Uniforms", SURFACE),
              ("Coaching", (180, 90, 90)), ("Other", TEXT_MUTED)]
    lx = dcx + rr + 30
    ly = dcy - rr + 6
    for i, (label, col) in enumerate(legend):
        yy = ly + i * 24
        d.rounded_rectangle((lx, yy, lx + 16, yy + 16), radius=3, fill=col)
        d.text((lx + 26, yy + 8), label, font=f(16), fill=TEXT, anchor="lm")

    # Bar: Budget vs Actual (top 6 categories)
    bbox = (gx + chart_w + gap, ch_y0, gx + 2 * chart_w + gap, ch_y0 + ch_h)
    d.rounded_rectangle(bbox, radius=10, fill=WHITE, outline=(220, 215, 205), width=2)
    d.text((bbox[0] + 16, bbox[1] + 14), "BUDGET vs ACTUAL",
           font=f(20), fill=ACCENT, anchor="lt")
    cats = ["Tourn", "Hotels", "Reg", "Coach", "Unif", "Fuel", "Food"]
    plan = [450, 380, 250, 300, 180, 240, 180]
    actual = [300, 280, 250, 150, 165, 188, 130]
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
    sbox = (cx - w // 2, cy - h // 2 + 30, cx + w // 2, cy + h // 2 + 80)
    shadow = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    ImageDraw.Draw(shadow).rounded_rectangle(sbox, radius=40, fill=(27, 79, 72, 110))
    canvas.alpha_composite(shadow.filter(ImageFilter.GaussianBlur(50)))
    body = (cx - w // 2, cy - h // 2, cx + w // 2, cy + h // 2)
    d.rounded_rectangle(body, radius=28, fill=(40, 40, 45), outline=(20, 20, 25), width=4)
    bezel = 22
    screen = (body[0] + bezel, body[1] + bezel, body[2] - bezel, body[3] - bezel)
    d.rounded_rectangle(screen, radius=8, fill=BG)
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

    pill(img, SIZE // 2, 230, "ULTIMATE 14-DASHBOARD SYSTEM",
         font=f(46), pad_x=70, pad_y=28)

    d = ImageDraw.Draw(img)
    text_center(d, (SIZE // 2, 430), "SOCCER MOM", f(190), PRIMARY)
    text_center(d, (SIZE // 2, 600), "COMMAND CENTER", f(160), PRIMARY)

    # Subtitle
    sub_font = f(48)
    parts = ["GAMES", "PRACTICE", "TEAM LIFE"]
    full = "  •  ".join(parts)
    tw, _ = text_size(d, full, sub_font)
    text_center(d, (SIZE // 2, 730), full, sub_font, TEXT)
    for sx in (SIZE // 2 - tw // 2 - 70, SIZE // 2 + tw // 2 + 70):
        d.polygon([(sx, 710), (sx + 20, 730), (sx, 750), (sx - 20, 730)], fill=PRIMARY)

    # Laptop with dashboard
    draw_laptop(img, SIZE // 2, 1240, w=1640, h=940)

    # Two floating KPI callouts
    kpi_callout(img, 360, 1010, "NEXT GAME", "2 DAYS",
                width=680, height=260, value_color=DANGER)
    kpi_callout(img, SIZE - 360, 1010, "BUDGET LEFT", "$214",
                width=680, height=260)

    # Bottom CTA
    pill(img, SIZE // 2, SIZE - 170,
         "GET THE 14-DASHBOARD SYSTEM",
         font=f(58), pad_x=80, pad_y=34, star=True)

    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
# Image 2: DASHBOARD CLOSEUP
# ===========================================================================
def render_dashboard_closeup(out: str) -> None:
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)

    pill(img, SIZE // 2, 180, "FAMILY COMMAND DASHBOARD",
         font=f(44), pad_x=60, pad_y=24)

    d = ImageDraw.Draw(img)
    text_center(d, (SIZE // 2, 320), "8 KPIs · 2 CHARTS · ZERO CHAOS", f(72), PRIMARY)
    text_center(d, (SIZE // 2, 410),
                "Games, practices, tournaments, budget — all live, all one screen",
                f(32, bold=False), TEXT_MUTED)

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
# Image 3: SEASON SCHEDULE
# ===========================================================================
def render_schedule(out: str) -> None:
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)
    d = ImageDraw.Draw(img)

    pill(img, SIZE // 2, 180, "SEASON SCHEDULE",
         font=f(44), pad_x=60, pad_y=24)
    text_center(d, (SIZE // 2, 320), "EVERY MATCH. COUNTDOWN INCLUDED.", f(64), PRIMARY)
    text_center(d, (SIZE // 2, 410),
                "Days-Out column auto-counts. Past games mute. Upcoming highlight mint.",
                f(30, bold=False), TEXT_MUTED)

    panel = (100, 510, SIZE - 100, SIZE - 180)
    drop_shadow(img, panel, radius=24, blur=40, alpha=110)
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    od.rounded_rectangle(panel, radius=24, fill=WHITE, outline=(220, 215, 205), width=3)
    img.alpha_composite(overlay)

    headers = ["DATE", "OPPONENT", "VENUE", "FIELD", "KICKOFF", "UNIFORM", "RESULT", "DAYS OUT"]
    rows = [
        ("Aug 24", "Riverside Rovers", "Home",    "Riverside Park",     "10:00", "Home Green",   "WIN  3-1", -28, "muted"),
        ("Aug 31", "Lincoln Lions",    "Away",    "Lincoln Complex",    "11:30", "Away White",   "DRAW 1-1", -21, "muted"),
        ("Sep 04", "Eastside Eagles",  "Away",    "Eastside Sports",    "9:00",  "Away White",   "LOSS 0-2", -17, "muted"),
        ("Sep 09", "Memorial Hawks",   "Home",    "Memorial Field",     "10:00", "Home Green",   "—",         2, "upcoming"),
        ("Sep 12", "Northside FC",     "Away",    "Northside Park",     "1:30",  "Away White",   "—",         5, "upcoming"),
        ("Sep 16", "Westside Wolves",  "Home",    "Riverside Park",     "11:00", "Home Green",   "—",         9, "soon"),
        ("Sep 19", "City Cup · Day 1", "Neutral", "Lincoln Complex",    "9:00",  "Home Green",   "—",        12, "tournament"),
        ("Sep 23", "City Cup · Day 2", "Neutral", "Lincoln Complex",    "1:00",  "Away White",   "—",        16, "tournament"),
        ("Sep 27", "Riverside Rovers", "Away",    "Riverside Park",     "10:00", "Away White",   "—",        20, "normal"),
        ("Oct 04", "Memorial Hawks",   "Away",    "Memorial Field",     "11:30", "Away White",   "—",        27, "normal"),
    ]
    inner = (panel[0] + 20, panel[1] + 20, panel[2] - 20, panel[3] - 20)
    iw = inner[2] - inner[0]
    ih = inner[3] - inner[1]
    cols = len(headers)
    col_w = iw / cols
    hdr_h = 80

    overlay2 = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay2)
    od.rounded_rectangle((inner[0], inner[1], inner[2], inner[1] + hdr_h),
                         radius=14, fill=PRIMARY)
    od.rectangle((inner[0], inner[1] + 30, inner[2], inner[1] + hdr_h), fill=PRIMARY)
    for i, h in enumerate(headers):
        cx = inner[0] + col_w * (i + 0.5)
        od.text((cx, inner[1] + hdr_h // 2), h, font=f(24), fill=WHITE, anchor="mm")

    row_h = (ih - hdr_h - 12) / len(rows)
    for ri, row in enumerate(rows):
        y0 = inner[1] + hdr_h + ri * row_h
        state = row[-1]
        cells = row[:-1]
        bg_fill = BG
        if state == "muted":
            bg_fill = (241, 241, 241)
        elif state == "upcoming":
            bg_fill = MINT_BG
        elif state == "soon":
            bg_fill = MINT_BG
        elif state == "tournament":
            bg_fill = WARN_BG
        elif ri % 2 == 1:
            bg_fill = (244, 236, 222)
        if bg_fill != BG:
            od.rectangle((inner[0] + 4, y0, inner[2] - 4, y0 + row_h), fill=bg_fill)
        for ci, val in enumerate(cells):
            cx = inner[0] + col_w * (ci + 0.5)
            color = PRIMARY if ci == 0 else TEXT
            if ci == 6:  # Result
                val_str = str(val)
                if "WIN" in val_str:
                    color = (40, 130, 100)
                elif "LOSS" in val_str:
                    color = DANGER
                elif "DRAW" in val_str:
                    color = ACCENT
            elif ci == 7:  # Days Out
                days = int(val)
                if days < 0:
                    color = TEXT_MUTED
                    val = f"{days}"
                elif days == 0:
                    color = DANGER
                    val = "TODAY"
                elif days <= 7:
                    color = DANGER
                    val = f"+{days}d"
                else:
                    color = PRIMARY
                    val = f"+{days}d"
            font_use = f(22) if ci in (0, 6, 7) else f(20)
            od.text((cx, y0 + row_h / 2), str(val), font=font_use, fill=color, anchor="mm")

    img.alpha_composite(overlay2)
    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
# Image 4: BUDGET
# ===========================================================================
def render_budget(out: str) -> None:
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)
    d = ImageDraw.Draw(img)

    pill(img, SIZE // 2, 180, "SOCCER BUDGET",
         font=f(44), pad_x=60, pad_y=24)
    text_center(d, (SIZE // 2, 320), "TRACK EVERY DOLLAR.", f(76), PRIMARY)
    text_center(d, (SIZE // 2, 410),
                "Per-month · per-player cost · auto variance · live charts",
                f(32, bold=False), TEXT_MUTED)

    # KPI strip
    kpis = [("PLANNED", "$2,470"), ("ACTUAL", "$2,036"),
            ("VARIANCE", "+$434"), ("PER PLAYER", "$1,018")]
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
        val_color = HIGHLIGHT if lab == "VARIANCE" else PRIMARY
        if lab == "VARIANCE":
            val_color = (40, 130, 100)
        od.text((x + kw // 2, ky + 120), val, font=f(64), fill=val_color, anchor="mm")
        img.alpha_composite(overlay)

    # Donut card
    dbox = (140, 760, 950, SIZE - 200)
    drop_shadow(img, dbox, radius=24, blur=30, alpha=90)
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    od.rounded_rectangle(dbox, radius=24, fill=WHITE, outline=(220, 215, 205), width=2)
    od.text((dbox[0] + 30, dbox[1] + 30), "SPENDING BREAKDOWN",
            font=f(30), fill=ACCENT, anchor="lt")
    dcx = (dbox[0] + dbox[2]) // 2 - 100
    dcy = (dbox[1] + dbox[3]) // 2 + 30
    rr = 220
    segs = [
        (22, PRIMARY,   "Tournaments 22%"),
        (16, ACCENT,    "Hotels 16%"),
        (14, HIGHLIGHT, "Registration 14%"),
        (10, SURFACE,   "Uniforms 10%"),
        (10, (180, 90, 90), "Coaching 10%"),
        (10, (200, 170, 110), "Fuel 10%"),
        (18, TEXT_MUTED, "Other 18%"),
    ]
    angle = -90
    for pct, col, _ in segs:
        sweep = pct * 3.6
        od.pieslice((dcx - rr, dcy - rr, dcx + rr, dcy + rr),
                    angle, angle + sweep, fill=col)
        angle += sweep
    od.ellipse((dcx - 90, dcy - 90, dcx + 90, dcy + 90), fill=WHITE)
    od.text((dcx, dcy - 12), "$2,036", font=f(36), fill=PRIMARY, anchor="mm")
    od.text((dcx, dcy + 30), "SPENT", font=f(20), fill=TEXT_MUTED, anchor="mm")
    # Legend
    lx = dcx + rr + 40
    for i, (pct, col, label) in enumerate(segs):
        yy = dcy - rr + 16 + i * 46
        od.rounded_rectangle((lx, yy, lx + 28, yy + 28), radius=5, fill=col)
        od.text((lx + 44, yy + 14), label, font=f(22), fill=TEXT, anchor="lm")
    img.alpha_composite(overlay)

    # Bar chart card
    bbox = (1010, 760, SIZE - 140, SIZE - 200)
    drop_shadow(img, bbox, radius=24, blur=30, alpha=90)
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    od.rounded_rectangle(bbox, radius=24, fill=WHITE, outline=(220, 215, 205), width=2)
    od.text((bbox[0] + 30, bbox[1] + 30), "BUDGET vs ACTUAL",
            font=f(30), fill=ACCENT, anchor="lt")
    cats = ["Reg", "Unif", "Cleats", "Balls", "Tourn", "Hotels", "Fuel", "Coach"]
    plan = [250, 180, 95, 45, 450, 380, 240, 300]
    actual = [250, 165, 90, 42, 300, 280, 188, 150]
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
    od.rounded_rectangle((bbox[0] + 30, bbox[3] - 50, bbox[0] + 50, bbox[3] - 30),
                         radius=4, fill=SURFACE)
    od.text((bbox[0] + 60, bbox[3] - 40), "Planned", font=f(20), fill=TEXT, anchor="lm")
    od.rounded_rectangle((bbox[0] + 180, bbox[3] - 50, bbox[0] + 200, bbox[3] - 30),
                         radius=4, fill=PRIMARY)
    od.text((bbox[0] + 210, bbox[3] - 40), "Actual", font=f(20), fill=TEXT, anchor="lm")
    img.alpha_composite(overlay)

    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
# Image 5: TOURNAMENT + PACKING
# ===========================================================================
def render_tournament(out: str) -> None:
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)
    d = ImageDraw.Draw(img)

    pill(img, SIZE // 2, 180, "TOURNAMENT MODE",
         font=f(44), pad_x=60, pad_y=24)
    text_center(d, (SIZE // 2, 320), "PLAN. PACK. WIN.", f(82), PRIMARY)
    text_center(d, (SIZE // 2, 410),
                "Tournament planner + 6 pre-built packing lists — nothing forgotten",
                f(32, bold=False), TEXT_MUTED)

    # Tournament planner table (top half)
    tbox = (140, 500, SIZE - 140, 1180)
    drop_shadow(img, tbox, radius=24, blur=30, alpha=100)
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    od.rounded_rectangle(tbox, radius=24, fill=WHITE, outline=(220, 215, 205), width=2)
    od.text((tbox[0] + 30, tbox[1] + 24), "UPCOMING TOURNAMENTS",
            font=f(28), fill=ACCENT, anchor="lt")
    img.alpha_composite(overlay)

    headers = ["TOURNAMENT", "DATES", "VENUE", "HOTEL", "FEE", "MATCHES", "DAYS"]
    rows = [
        ("City Cup",             "Sep 19-23", "Lincoln Complex",  "Hampton Inn Lincoln",   "$125", "4", "+12"),
        ("Riverbend Classic",    "Oct 11-13", "Riverbend Sports", "Marriott Riverbend",    "$150", "3", "+34"),
        ("Fall Showcase",        "Nov 06-09", "Capitol Fields",   "Hyatt Place Capitol",   "$195", "5", "+60"),
        ("State Cup Qualifier",  "Dec 01-04", "State Complex",    "Holiday Inn State",     "$220", "3", "+85"),
    ]
    inner = (tbox[0] + 30, tbox[1] + 80, tbox[2] - 30, tbox[3] - 30)
    iw = inner[2] - inner[0]
    ih = inner[3] - inner[1]
    cols = len(headers)
    col_w = iw / cols
    hdr_h = 60

    overlay2 = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay2)
    od.rounded_rectangle((inner[0], inner[1], inner[2], inner[1] + hdr_h),
                         radius=10, fill=PRIMARY)
    od.rectangle((inner[0], inner[1] + 20, inner[2], inner[1] + hdr_h), fill=PRIMARY)
    for i, h in enumerate(headers):
        cx = inner[0] + col_w * (i + 0.5)
        od.text((cx, inner[1] + hdr_h // 2), h, font=f(22), fill=WHITE, anchor="mm")

    row_h = (ih - hdr_h - 8) / len(rows)
    for ri, row in enumerate(rows):
        y0 = inner[1] + hdr_h + ri * row_h
        if ri == 0:
            od.rectangle((inner[0] + 2, y0, inner[2] - 2, y0 + row_h), fill=MINT_BG)
        elif ri % 2 == 1:
            od.rectangle((inner[0] + 2, y0, inner[2] - 2, y0 + row_h), fill=(244, 236, 222))
        for ci, val in enumerate(row):
            cx = inner[0] + col_w * (ci + 0.5)
            color = PRIMARY if ci in (0, 6) else TEXT
            od.text((cx, y0 + row_h / 2), str(val), font=f(22), fill=color, anchor="mm")
    img.alpha_composite(overlay2)

    # Packing checklist (bottom half) - 3 columns
    pbox = (140, 1220, SIZE - 140, SIZE - 130)
    drop_shadow(img, pbox, radius=24, blur=30, alpha=100)
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    od.rounded_rectangle(pbox, radius=24, fill=WHITE, outline=(220, 215, 205), width=2)
    od.text((pbox[0] + 30, pbox[1] + 24), "TOURNAMENT WEEKEND CHECKLIST",
            font=f(28), fill=ACCENT, anchor="lt")
    img.alpha_composite(overlay)

    cols_data = [
        [("Match cleats", True),
         ("Spare cleats", True),
         ("Shin guards", True),
         ("Match socks x3", True),
         ("Home + away kits", True),
         ("GK gloves", False),
         ("Snack + electrolytes", True)],
        [("Water bottles (filled)", True),
         ("Spare ball", True),
         ("Pop-up tent / shade", True),
         ("Cooler with ice", True),
         ("Camp chair", False),
         ("Sunscreen", True),
         ("First-aid kit", True)],
        [("Hotel essentials", True),
         ("Phone charger", True),
         ("Pajamas", True),
         ("Game schedule", True),
         ("Hotel confirmation", True),
         ("Driver list", True),
         ("Cash for entry", False)],
    ]
    inner = (pbox[0] + 30, pbox[1] + 90, pbox[2] - 30, pbox[3] - 30)
    cw = (inner[2] - inner[0]) / 3
    overlay3 = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay3)
    for ci, items in enumerate(cols_data):
        col_x = inner[0] + ci * cw
        for ri, (item, checked) in enumerate(items):
            y = inner[1] + ri * 60
            cbox = (col_x + 10, y, col_x + 50, y + 40)
            if checked:
                od.rounded_rectangle(cbox, radius=8, fill=PRIMARY)
                od.text((col_x + 30, y + 20), "✓", font=f(28), fill=WHITE, anchor="mm")
            else:
                od.rounded_rectangle(cbox, radius=8, fill=BG,
                                     outline=PRIMARY, width=3)
            text_color = PRIMARY if checked else TEXT
            od.text((col_x + 70, y + 20), item, font=f(24), fill=text_color, anchor="lm")
    img.alpha_composite(overlay3)

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
                "Check the schedule sideline-side, update on the drive home",
                f(32, bold=False), TEXT_MUTED)

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
    od.rounded_rectangle((sx0, sy0, sx1, sy0 + 90), radius=42, fill=PRIMARY)
    od.rectangle((sx0, sy0 + 50, sx1, sy0 + 90), fill=PRIMARY)
    od.text(((sx0 + sx1) // 2, sy0 + 50), "Command Center",
            font=f(32), fill=WHITE, anchor="mm")

    # KPI cards stacked
    y = sy0 + 130
    kpis_mobile = [
        ("NEXT GAME",      "2 DAYS",   DANGER),
        ("BUDGET REMAINING","$214",   PRIMARY),
        ("ATTENDANCE %",    "92%",     PRIMARY),
        ("EQUIPMENT NEEDED","1 item",  ACCENT),
    ]
    for lab, val, col in kpis_mobile:
        cbox = (sx0 + 30, y, sx1 - 30, y + 140)
        od.rounded_rectangle(cbox, radius=18, fill=WHITE, outline=(220, 215, 205), width=2)
        od.text((cbox[0] + 26, y + 30), lab, font=f(22), fill=ACCENT, anchor="lt")
        od.text((cbox[0] + 26, y + 96), val, font=f(46), fill=col, anchor="lm")
        y += 160

    # Upcoming games list
    od.text((sx0 + 40, y + 20), "UPCOMING", font=f(22), fill=ACCENT, anchor="lt")
    y += 50
    events = [
        ("Sep 09", "vs Memorial Hawks",  HIGHLIGHT, "+2d"),
        ("Sep 12", "@ Northside FC",     SURFACE,   "+5d"),
        ("Sep 16", "vs Westside Wolves", SURFACE,   "+9d"),
        ("Sep 19", "City Cup · Day 1",   WARN_BG,   "+12d"),
    ]
    for date_str, evt, dot_col, days in events:
        od.ellipse((sx0 + 40, y + 14, sx0 + 64, y + 38), fill=dot_col)
        od.text((sx0 + 80, y + 14), date_str, font=f(20), fill=PRIMARY, anchor="lt")
        od.text((sx0 + 80, y + 40), evt, font=f(22), fill=TEXT, anchor="lt")
        od.text((sx1 - 50, y + 30), days, font=f(22), fill=PRIMARY, anchor="rm")
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
        ("01_hero.png",       render_hero),
        ("02_dashboard.png",  render_dashboard_closeup),
        ("03_schedule.png",   render_schedule),
        ("04_budget.png",     render_budget),
        ("05_tournament.png", render_tournament),
        ("06_mobile.png",     render_mobile),
    ]
    for name, fn in targets:
        path = os.path.join(out_dir, name)
        fn(path)
        print(f"  ✓ {name}")
    print(f"Wrote {len(targets)} images to {out_dir}")


if __name__ == "__main__":
    main()

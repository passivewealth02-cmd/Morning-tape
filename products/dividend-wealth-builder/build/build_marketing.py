"""Generate the Etsy marketing image set for Dividend Wealth Builder.

Outputs six 2000x2000 PNGs to ../marketing/:
  01_hero.png            - main listing thumbnail (Driver-Budget-style layout)
  02_dashboard.png       - dashboard close-up
  03_holdings.png        - holdings table close-up
  04_charts.png          - charts grid
  05_projections.png     - FIRE projection highlight
  06_mobile.png          - mobile Google Sheets preview

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
PRIMARY = (27, 79, 72)        # #1B4F48
PRIMARY_DARK = (15, 50, 45)
ACCENT = (147, 115, 86)       # #937356
SURFACE = (229, 211, 186)     # #E5D3BA
HIGHLIGHT = (117, 230, 193)   # #75E6C1
BG = (252, 250, 246)          # near-white off-cream
WHITE = (255, 255, 255)
TEXT = (51, 51, 51)
TEXT_MUTED = (120, 120, 120)
DANGER = (201, 76, 76)
DOT = (220, 215, 205)
SHADOW = (27, 79, 72, 40)

FONT_BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
FONT_REG = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
FONT_MONO_BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSansMono-Bold.ttf"

SIZE = 2000


def f(size: int, bold: bool = True, mono: bool = False) -> ImageFont.FreeTypeFont:
    if mono:
        path = FONT_MONO_BOLD
    else:
        path = FONT_BOLD if bold else FONT_REG
    return ImageFont.truetype(path, size)


# ---------------------------------------------------------------------------
# Primitives
# ---------------------------------------------------------------------------
def dotted_bg(canvas: Image.Image, dot_color=DOT, spacing: int = 32, radius: int = 3) -> None:
    d = ImageDraw.Draw(canvas)
    for y in range(spacing // 2, canvas.height, spacing):
        for x in range(spacing // 2, canvas.width, spacing):
            d.ellipse((x - radius, y - radius, x + radius, y + radius), fill=dot_color)


def rounded_rect(
    draw: ImageDraw.ImageDraw, box, radius: int, fill=None, outline=None, width: int = 0
):
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


def drop_shadow(canvas: Image.Image, box, radius: int, blur: int = 24, alpha: int = 70):
    """Soft drop shadow under a rounded rect at `box`."""
    shadow_layer = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    sd = ImageDraw.Draw(shadow_layer)
    sd.rounded_rectangle(box, radius=radius, fill=(27, 79, 72, alpha))
    shadow_layer = shadow_layer.filter(ImageFilter.GaussianBlur(blur))
    canvas.alpha_composite(shadow_layer, (0, 18))


def text_center(draw, xy, text, font, fill, anchor="mm"):
    draw.text(xy, text, font=font, fill=fill, anchor=anchor)


def text_size(draw, text, font):
    box = draw.textbbox((0, 0), text, font=font)
    return box[2] - box[0], box[3] - box[1]


def pill(
    canvas: Image.Image,
    cx: int,
    cy: int,
    text: str,
    font,
    fg=WHITE,
    bg=PRIMARY,
    pad_x: int = 60,
    pad_y: int = 26,
    border: tuple | None = None,
    border_width: int = 0,
    star: bool = False,
):
    """Draw a centered rounded pill at (cx, cy)."""
    overlay = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(overlay)
    label = f"★ {text}" if star else text
    tw, th = text_size(d, label, font)
    w = tw + pad_x * 2
    h = th + pad_y * 2
    box = (cx - w // 2, cy - h // 2, cx + w // 2, cy + h // 2)
    r = h // 2
    if border:
        d.rounded_rectangle(box, radius=r, fill=bg, outline=border, width=border_width)
    else:
        d.rounded_rectangle(box, radius=r, fill=bg)
    d.text((cx, cy), label, font=font, fill=fg, anchor="mm")
    canvas.alpha_composite(overlay)
    return box


def kpi_callout(
    canvas: Image.Image,
    cx: int,
    cy: int,
    label: str,
    value: str,
    width: int = 720,
    height: int = 280,
):
    """Big floating white KPI card (used over the laptop)."""
    box = (cx - width // 2, cy - height // 2, cx + width // 2, cy + height // 2)
    drop_shadow(canvas, box, radius=24, blur=30, alpha=85)
    overlay = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(overlay)
    d.rounded_rectangle(box, radius=24, fill=WHITE, outline=(220, 215, 205), width=2)
    d.text((cx, cy - 70), label, font=f(46), fill=TEXT, anchor="mm")
    d.text((cx, cy + 50), value, font=f(108), fill=PRIMARY, anchor="mm")
    canvas.alpha_composite(overlay)


# ---------------------------------------------------------------------------
# Laptop + dashboard rendering
# ---------------------------------------------------------------------------
def draw_laptop(canvas: Image.Image, cx: int, cy: int, w: int = 1500, h: int = 900):
    """Draw a stylised laptop with a brand-themed dashboard on the screen."""
    overlay = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(overlay)

    # Soft shadow under laptop
    sbox = (cx - w // 2, cy - h // 2 + 30, cx + w // 2, cy + h // 2 + 80)
    shadow_layer = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    sd = ImageDraw.Draw(shadow_layer)
    sd.rounded_rectangle(sbox, radius=40, fill=(27, 79, 72, 110))
    shadow_layer = shadow_layer.filter(ImageFilter.GaussianBlur(50))
    canvas.alpha_composite(shadow_layer)

    # Laptop body (bezel)
    body = (cx - w // 2, cy - h // 2, cx + w // 2, cy + h // 2)
    d.rounded_rectangle(body, radius=28, fill=(40, 40, 45), outline=(20, 20, 25), width=4)

    # Screen area
    bezel = 22
    screen = (body[0] + bezel, body[1] + bezel, body[2] - bezel, body[3] - bezel)
    d.rounded_rectangle(screen, radius=8, fill=BG)

    # Bottom base (hinge bar suggestion)
    base_h = 24
    base = (cx - w // 2 - 120, cy + h // 2, cx + w // 2 + 120, cy + h // 2 + base_h)
    d.rounded_rectangle(base, radius=12, fill=(60, 60, 65))
    # Hinge slot
    d.rounded_rectangle(
        (cx - 80, cy + h // 2 + 4, cx + 80, cy + h // 2 + 14),
        radius=6,
        fill=(35, 35, 40),
    )

    canvas.alpha_composite(overlay)

    # Now draw dashboard contents directly onto canvas using screen coords
    draw_dashboard_in_screen(canvas, screen)


def draw_dashboard_in_screen(canvas: Image.Image, screen):
    sx0, sy0, sx1, sy1 = screen
    sw = sx1 - sx0
    sh = sy1 - sy0
    overlay = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(overlay)

    # Top header bar
    head_h = 86
    d.rounded_rectangle(
        (sx0, sy0, sx1, sy0 + head_h), radius=8, fill=PRIMARY
    )
    d.rectangle((sx0, sy0 + 30, sx1, sy0 + head_h), fill=PRIMARY)
    d.text((sx0 + 28, sy0 + head_h // 2), "DASHBOARD OVERVIEW",
           font=f(28), fill=WHITE, anchor="lm")
    d.text((sx1 - 28, sy0 + head_h // 2), "SELECT MONTH ▾   May 2024",
           font=f(22), fill=WHITE, anchor="rm")

    # KPI strip
    kpi_y = sy0 + head_h + 28
    kpi_h = 130
    gutter = 22
    kpi_w = (sw - 28 * 2 - gutter * 3) // 4
    kpi_labels = [
        ("PORTFOLIO VALUE", "$148,420"),
        ("MONTHLY INCOME", "$612.40"),
        ("ANNUAL INCOME", "$7,349"),
        ("YIELD %", "4.95%"),
    ]
    for i, (lab, val) in enumerate(kpi_labels):
        x0 = sx0 + 28 + i * (kpi_w + gutter)
        box = (x0, kpi_y, x0 + kpi_w, kpi_y + kpi_h)
        d.rounded_rectangle(box, radius=10, fill=WHITE, outline=(220, 215, 205), width=2)
        d.text((x0 + 16, kpi_y + 18), lab, font=f(20), fill=ACCENT, anchor="lt")
        d.text((x0 + 16, kpi_y + 72), val, font=f(46), fill=PRIMARY, anchor="lm")

    # Two-chart row
    cy0 = kpi_y + kpi_h + 30
    cw = (sw - 28 * 2 - gutter) // 2
    ch = 270

    # Bar chart (left)
    bar_box = (sx0 + 28, cy0, sx0 + 28 + cw, cy0 + ch)
    d.rounded_rectangle(bar_box, radius=10, fill=WHITE, outline=(220, 215, 205), width=2)
    d.text((bar_box[0] + 16, bar_box[1] + 14), "MONTHLY INCOME",
           font=f(20), fill=ACCENT, anchor="lt")
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"]
    heights = [80, 95, 110, 140, 160, 180]
    bx = bar_box[0] + 40
    by = bar_box[3] - 30
    bar_w = 36
    bar_gap = 32
    for i, h in enumerate(heights):
        x = bx + i * (bar_w + bar_gap)
        d.rounded_rectangle((x, by - h, x + bar_w, by), radius=4, fill=PRIMARY)
        d.text((x + bar_w // 2, by + 14), months[i], font=f(18), fill=TEXT_MUTED, anchor="mt")

    # Line chart (right)
    line_box = (sx0 + 28 + cw + gutter, cy0, sx0 + 28 + cw + gutter + cw, cy0 + ch)
    d.rounded_rectangle(line_box, radius=10, fill=WHITE, outline=(220, 215, 205), width=2)
    d.text((line_box[0] + 16, line_box[1] + 14), "WEALTH GROWTH",
           font=f(20), fill=ACCENT, anchor="lt")
    lx0 = line_box[0] + 40
    ly0 = line_box[1] + 60
    lx1 = line_box[2] - 30
    ly1 = line_box[3] - 30
    pts = []
    n = 14
    for i in range(n):
        t = i / (n - 1)
        v = (1.0 - math.exp(-2.0 * t)) * 0.95 + t * 0.05
        x = lx0 + t * (lx1 - lx0)
        y = ly1 - v * (ly1 - ly0)
        pts.append((x, y))
    # Fill under curve
    poly = pts + [(lx1, ly1), (lx0, ly1)]
    d.polygon(poly, fill=(117, 230, 193, 90))
    d.line(pts, fill=PRIMARY, width=4)
    for p in pts[::2]:
        d.ellipse((p[0] - 5, p[1] - 5, p[0] + 5, p[1] + 5), fill=PRIMARY)

    # Donut (over line chart, top-right corner)
    donut_cx = line_box[2] - 70
    donut_cy = line_box[1] + 76
    rr = 36
    d.pieslice((donut_cx - rr, donut_cy - rr, donut_cx + rr, donut_cy + rr),
               -90, 60, fill=PRIMARY)
    d.pieslice((donut_cx - rr, donut_cy - rr, donut_cx + rr, donut_cy + rr),
               60, 180, fill=ACCENT)
    d.pieslice((donut_cx - rr, donut_cy - rr, donut_cx + rr, donut_cy + rr),
               180, 270, fill=HIGHLIGHT)
    d.ellipse((donut_cx - 18, donut_cy - 18, donut_cx + 18, donut_cy + 18), fill=WHITE)

    # Data table at bottom
    tbl_y = cy0 + ch + 26
    tbl_h = sy1 - tbl_y - 28
    tbl_box = (sx0 + 28, tbl_y, sx1 - 28, tbl_y + tbl_h)
    d.rounded_rectangle(tbl_box, radius=10, fill=WHITE, outline=(220, 215, 205), width=2)
    headers = ["TICKER", "SECTOR", "SHARES", "VALUE", "INCOME", "YIELD"]
    rows = [
        ("AAPL",  "Tech",       "50",  "$9,620", "$192",   "1.99%"),
        ("MSFT",  "Tech",       "40",  "$16,608","$480",   "2.89%"),
        ("O",     "REITs",      "200", "$11,620","$2,528", "5.44%"),
        ("JPM",   "Finance",    "35",  "$6,933", "$644",   "9.29%"),
        ("KO",    "Consumer",   "120", "$7,608", "$883",   "11.61%"),
    ]
    # Header bar
    hdr_h = 36
    d.rounded_rectangle((tbl_box[0], tbl_box[1], tbl_box[2], tbl_box[1] + hdr_h),
                        radius=10, fill=PRIMARY)
    d.rectangle((tbl_box[0], tbl_box[1] + 18, tbl_box[2], tbl_box[1] + hdr_h),
                fill=PRIMARY)
    cols = 6
    col_w = (tbl_box[2] - tbl_box[0]) / cols
    for i, hh in enumerate(headers):
        cx = tbl_box[0] + col_w * (i + 0.5)
        d.text((cx, tbl_box[1] + hdr_h // 2), hh, font=f(18), fill=WHITE, anchor="mm")
    row_h = (tbl_h - hdr_h - 8) / len(rows)
    for ri, row in enumerate(rows):
        y0 = tbl_box[1] + hdr_h + ri * row_h
        if ri % 2 == 1:
            d.rectangle((tbl_box[0] + 2, y0, tbl_box[2] - 2, y0 + row_h),
                        fill=(244, 236, 222))
        for ci, val in enumerate(row):
            cx = tbl_box[0] + col_w * (ci + 0.5)
            color = PRIMARY if ci == 0 else TEXT
            d.text((cx, y0 + row_h // 2), val, font=f(20), fill=color, anchor="mm")

    canvas.alpha_composite(overlay)


# ===========================================================================
# Image 1: HERO
# ===========================================================================
def render_hero(out: str) -> None:
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img, spacing=44, radius=4)

    # Top badge pill
    pill(img, SIZE // 2, 230, "ULTIMATE 5-DASHBOARD SYSTEM",
         font=f(46), bg=PRIMARY, pad_x=70, pad_y=28)

    # Wordmark (two-line dominant title)
    d = ImageDraw.Draw(img)
    title = "DIVIDEND WEALTH"
    tfont = f(180)
    text_center(d, (SIZE // 2, 430), title, tfont, PRIMARY)
    text_center(d, (SIZE // 2, 590), "BUILDER", f(180), PRIMARY)

    # Subtitle row with dots and small diamond icons
    subline_y = 720
    sub_font = f(50)
    parts = ["DIVIDENDS", "FIRE", "PASSIVE INCOME"]
    sep = "  •  "
    full = sep.join(parts)
    tw, _ = text_size(d, full, sub_font)
    text_center(d, (SIZE // 2, subline_y), full, sub_font, TEXT)
    # Small primary diamonds either side
    for sx in (SIZE // 2 - tw // 2 - 70, SIZE // 2 + tw // 2 + 70):
        d.polygon([(sx, subline_y - 20), (sx + 20, subline_y),
                   (sx, subline_y + 20), (sx - 20, subline_y)], fill=PRIMARY)

    # Laptop with dashboard
    draw_laptop(img, SIZE // 2, 1240, w=1640, h=940)

    # Two big floating KPI callout cards over laptop
    kpi_callout(img, 360, 1010, "MONTHLY INCOME", "$612.40", width=680, height=260)
    kpi_callout(img, SIZE - 360, 1010, "PORTFOLIO VALUE", "$148,420", width=680, height=260)

    # Bottom CTA pill
    pill(img, SIZE // 2, SIZE - 170,
         "GET THE 5-DASHBOARD SYSTEM",
         font=f(58), bg=PRIMARY, pad_x=80, pad_y=34, star=True)

    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
# Image 2: DASHBOARD CLOSEUP
# ===========================================================================
def render_dashboard_closeup(out: str) -> None:
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img, spacing=44, radius=4)

    pill(img, SIZE // 2, 180, "LIVE DASHBOARD PREVIEW",
         font=f(44), bg=PRIMARY, pad_x=60, pad_y=24)

    d = ImageDraw.Draw(img)
    text_center(d, (SIZE // 2, 320), "EVERY METRIC, ONE SCREEN", f(86), PRIMARY)
    text_center(d, (SIZE // 2, 410),
                "Track portfolio value, yield, monthly income & progress to FIRE",
                f(36, bold=False), TEXT_MUTED)

    # Big dashboard panel (no laptop frame — close-up)
    panel_box = (140, 500, SIZE - 140, SIZE - 200)
    drop_shadow(img, panel_box, radius=24, blur=40, alpha=110)
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    od.rounded_rectangle(panel_box, radius=24, fill=BG, outline=(220, 215, 205), width=3)
    img.alpha_composite(overlay)
    inner = (panel_box[0] + 24, panel_box[1] + 24, panel_box[2] - 24, panel_box[3] - 24)
    draw_dashboard_in_screen(img, inner)

    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
# Image 3: HOLDINGS TABLE
# ===========================================================================
def render_holdings(out: str) -> None:
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img, spacing=44, radius=4)
    d = ImageDraw.Draw(img)

    pill(img, SIZE // 2, 180, "HOLDINGS DATABASE",
         font=f(44), bg=PRIMARY, pad_x=60, pad_y=24)
    text_center(d, (SIZE // 2, 320), "VALIDATED · FORMATTED · LIVE",
                f(78), PRIMARY)
    text_center(d, (SIZE // 2, 410),
                "Drop in tickers — formulas, dropdowns & risk bars do the rest",
                f(34, bold=False), TEXT_MUTED)

    # Big table panel
    panel = (120, 510, SIZE - 120, SIZE - 180)
    drop_shadow(img, panel, radius=24, blur=40, alpha=110)
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    od.rounded_rectangle(panel, radius=24, fill=WHITE, outline=(220, 215, 205), width=3)
    img.alpha_composite(overlay)

    headers = ["TICKER", "COMPANY", "SECTOR", "SHARES",
               "VALUE", "DIV/SH", "INCOME", "YIELD", "RISK"]
    rows = [
        ("AAPL", "Apple Inc.",        "Tech",       "50",  "$9,620",  "$0.96", "$192",   "1.99%",  2),
        ("MSFT", "Microsoft Corp.",   "Tech",       "40",  "$16,608", "$3.00", "$480",   "2.89%",  2),
        ("JNJ",  "Johnson & Johnson", "Healthcare", "60",  "$9,492",  "$4.76", "$1,142", "12.03%", 1),
        ("KO",   "Coca-Cola Co.",     "Consumer",   "120", "$7,608",  "$1.84", "$883",   "11.61%", 1),
        ("PG",   "Procter & Gamble",  "Consumer",   "45",  "$7,281",  "$4.03", "$726",   "9.97%",  1),
        ("JPM",  "JPMorgan Chase",    "Finance",    "35",  "$6,933",  "$4.60", "$644",   "9.29%",  3),
        ("XOM",  "Exxon Mobil",       "Energy",     "80",  "$9,304",  "$3.80", "$1,216", "13.07%", 3),
        ("O",    "Realty Income",     "REITs",      "200", "$11,620", "$3.16", "$2,528", "21.75%", 2),
        ("VYM",  "Vanguard High Div", "Finance",    "100", "$12,480", "$3.45", "$345",   "2.76%",  2),
        ("MAIN", "Main St. Capital",  "Finance",    "90",  "$4,482",  "$2.94", "$3,175", "70.85%", 3),
    ]
    # Header
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
        od.text((cx, inner[1] + hdr_h // 2), h, font=f(28), fill=WHITE, anchor="mm")

    row_h = (ih - hdr_h - 12) / len(rows)
    for ri, row in enumerate(rows):
        y0 = inner[1] + hdr_h + ri * row_h
        if ri % 2 == 1:
            od.rectangle((inner[0] + 4, y0, inner[2] - 4, y0 + row_h),
                         fill=(244, 236, 222))
        for ci, val in enumerate(row[:-1]):
            cx = inner[0] + col_w * (ci + 0.5)
            color = PRIMARY if ci == 0 else TEXT
            od.text((cx, y0 + row_h // 2), str(val), font=f(28), fill=color, anchor="mm")
        # Risk data bar
        risk = row[-1]
        rcx = inner[0] + col_w * (cols - 0.5)
        bar_w = col_w * 0.7
        bar_x0 = rcx - bar_w / 2
        bar_y0 = y0 + row_h / 2 - 16
        bar_y1 = y0 + row_h / 2 + 16
        od.rounded_rectangle((bar_x0, bar_y0, bar_x0 + bar_w, bar_y1),
                             radius=6, fill=(245, 235, 235))
        fill_w = bar_w * (risk / 5)
        od.rounded_rectangle((bar_x0, bar_y0, bar_x0 + fill_w, bar_y1),
                             radius=6, fill=DANGER)
        od.text((rcx, y0 + row_h // 2), str(risk),
                font=f(24), fill=WHITE if risk >= 2 else PRIMARY, anchor="mm")

    img.alpha_composite(overlay2)
    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
# Image 4: CHARTS GRID
# ===========================================================================
def render_charts(out: str) -> None:
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img, spacing=44, radius=4)
    d = ImageDraw.Draw(img)

    pill(img, SIZE // 2, 180, "VISUAL INTELLIGENCE",
         font=f(44), bg=PRIMARY, pad_x=60, pad_y=24)
    text_center(d, (SIZE // 2, 320), "6 CHARTS · 1 GLANCE", f(86), PRIMARY)
    text_center(d, (SIZE // 2, 410),
                "Allocation, growth, dividend cadence — all auto-updating",
                f(34, bold=False), TEXT_MUTED)

    # Four chart cards in 2x2 grid
    grid_top = 540
    grid_left = 140
    cell_w = (SIZE - grid_left * 2 - 60) // 2
    cell_h = 600
    gap = 60

    def card(x, y, title):
        box = (x, y, x + cell_w, y + cell_h)
        drop_shadow(img, box, radius=20, blur=30, alpha=80)
        overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
        od = ImageDraw.Draw(overlay)
        od.rounded_rectangle(box, radius=20, fill=WHITE, outline=(220, 215, 205), width=2)
        od.text((x + 30, y + 30), title, font=f(32), fill=ACCENT, anchor="lt")
        img.alpha_composite(overlay)
        return box

    # 1. Sector pie (top-left)
    b = card(grid_left, grid_top, "ALLOCATION BY SECTOR")
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    pcx = b[0] + cell_w // 2
    pcy = b[1] + 340
    pr = 200
    segments = [
        (32, PRIMARY,  "Tech 32%"),
        (24, ACCENT,   "Finance 24%"),
        (18, HIGHLIGHT,"REITs 18%"),
        (12, SURFACE,  "Consumer 12%"),
        (8,  (200, 80, 80), "Energy 8%"),
        (6,  TEXT_MUTED, "Healthcare 6%"),
    ]
    angle = -90
    for pct, color, _ in segments:
        sweep = pct * 3.6
        od.pieslice((pcx - pr, pcy - pr, pcx + pr, pcy + pr),
                    angle, angle + sweep, fill=color)
        angle += sweep
    od.ellipse((pcx - 86, pcy - 86, pcx + 86, pcy + 86), fill=WHITE)
    od.text((pcx, pcy - 8), "$148K", font=f(40), fill=PRIMARY, anchor="mm")
    od.text((pcx, pcy + 32), "TOTAL", font=f(20), fill=TEXT_MUTED, anchor="mm")
    # Legend
    ly = b[3] - 130
    for i, (pct, color, label) in enumerate(segments):
        col = i % 3
        row = i // 3
        lx = b[0] + 40 + col * 220
        yy = ly + row * 44
        od.rounded_rectangle((lx, yy, lx + 22, yy + 22), radius=4, fill=color)
        od.text((lx + 34, yy + 11), label, font=f(20), fill=TEXT, anchor="lm")
    img.alpha_composite(overlay)

    # 2. Monthly income bar (top-right)
    b = card(grid_left + cell_w + gap, grid_top, "MONTHLY DIVIDEND INCOME")
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    vals = [410, 580, 450, 720, 510, 640, 480, 690, 520, 750, 580, 820]
    bx0 = b[0] + 60
    by0 = b[1] + 110
    bx1 = b[2] - 40
    by1 = b[3] - 90
    max_v = max(vals)
    bw = (bx1 - bx0) / (len(vals) * 1.5)
    for i, v in enumerate(vals):
        x = bx0 + i * (bw * 1.5)
        h = (v / max_v) * (by1 - by0)
        color = HIGHLIGHT if v == max_v else PRIMARY
        od.rounded_rectangle((x, by1 - h, x + bw, by1), radius=6, fill=color)
        od.text((x + bw / 2, by1 + 22), months[i],
                font=f(20), fill=TEXT_MUTED, anchor="mt")
    img.alpha_composite(overlay)

    # 3. Wealth growth (bottom-left)
    b = card(grid_left, grid_top + cell_h + gap, "WEALTH GROWTH PROJECTION")
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    lx0 = b[0] + 60
    ly0 = b[1] + 130
    lx1 = b[2] - 60
    ly1 = b[3] - 100
    pts = []
    n = 24
    for i in range(n):
        t = i / (n - 1)
        v = 0.1 + 0.9 * (1 - math.exp(-2.7 * t))
        pts.append((lx0 + t * (lx1 - lx0), ly1 - v * (ly1 - ly0)))
    poly = pts + [(lx1, ly1), (lx0, ly1)]
    od.polygon(poly, fill=(117, 230, 193, 110))
    od.line(pts, fill=PRIMARY, width=6)
    od.text((lx0, ly0 - 40), "$1.2M projected at year 30",
            font=f(26), fill=ACCENT, anchor="lt")
    # Axis hints
    for i, lbl in enumerate(["Y1", "Y10", "Y20", "Y30"]):
        x = lx0 + (i / 3) * (lx1 - lx0)
        od.text((x, ly1 + 22), lbl, font=f(20), fill=TEXT_MUTED, anchor="mt")
    img.alpha_composite(overlay)

    # 4. Yield vs Risk scatter (bottom-right)
    b = card(grid_left + cell_w + gap, grid_top + cell_h + gap, "YIELD vs RISK")
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    sx0 = b[0] + 80
    sy0 = b[1] + 130
    sx1 = b[2] - 60
    sy1 = b[3] - 100
    od.line((sx0, sy1, sx1, sy1), fill=(200, 200, 200), width=2)
    od.line((sx0, sy0, sx0, sy1), fill=(200, 200, 200), width=2)
    rnd = random.Random(7)
    for _ in range(28):
        rx = rnd.uniform(0.1, 0.95)
        ry = rnd.uniform(0.05, 0.9)
        x = sx0 + rx * (sx1 - sx0)
        y = sy1 - ry * (sy1 - sy0)
        r = 10 + rnd.randint(0, 10)
        color = HIGHLIGHT if ry > 0.55 else PRIMARY if ry > 0.3 else ACCENT
        od.ellipse((x - r, y - r, x + r, y + r), fill=color + (220,) if len(color) == 3 else color)
    od.text((sx0, sy1 + 14), "RISK →", font=f(20), fill=TEXT_MUTED, anchor="lt")
    od.text((sx0 - 14, sy0), "YIELD ↑", font=f(20), fill=TEXT_MUTED, anchor="rt")
    img.alpha_composite(overlay)

    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
# Image 5: PROJECTIONS / FIRE
# ===========================================================================
def render_projections(out: str) -> None:
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img, spacing=44, radius=4)
    d = ImageDraw.Draw(img)

    pill(img, SIZE // 2, 180, "FIRE PROJECTION ENGINE",
         font=f(44), bg=PRIMARY, pad_x=60, pad_y=24)
    text_center(d, (SIZE // 2, 320), "YOUR PATH TO", f(86), PRIMARY)
    text_center(d, (SIZE // 2, 430),
                "FINANCIAL INDEPENDENCE",
                f(86), PRIMARY)
    text_center(d, (SIZE // 2, 540),
                "Solve for years to $1k · $5k · custom target — automatically",
                f(34, bold=False), TEXT_MUTED)

    # 3 stat cards
    card_y = 660
    card_h = 320
    card_w = 540
    gap = 50
    total_w = card_w * 3 + gap * 2
    start_x = (SIZE - total_w) // 2

    stats = [
        ("YEARS TO $1K/MO", "4.2", "@ 8% portfolio growth"),
        ("YEARS TO $5K/MO", "14.8", "with monthly reinvest"),
        ("FIRE NUMBER", "$1.5M", "25× target spend"),
    ]
    for i, (lab, val, sub) in enumerate(stats):
        x = start_x + i * (card_w + gap)
        box = (x, card_y, x + card_w, card_y + card_h)
        drop_shadow(img, box, radius=24, blur=30, alpha=90)
        overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
        od = ImageDraw.Draw(overlay)
        od.rounded_rectangle(box, radius=24, fill=WHITE, outline=(220, 215, 205), width=2)
        od.text((x + card_w // 2, card_y + 60), lab,
                font=f(34), fill=ACCENT, anchor="mm")
        od.text((x + card_w // 2, card_y + 170), val,
                font=f(140), fill=PRIMARY, anchor="mm")
        od.text((x + card_w // 2, card_y + 270), sub,
                font=f(26, bold=False), fill=TEXT_MUTED, anchor="mm")
        img.alpha_composite(overlay)

    # Big curve at bottom
    cx0, cy0 = 200, 1050
    cx1, cy1 = SIZE - 200, SIZE - 180
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    box = (cx0, cy0, cx1, cy1)
    drop_shadow(img, box, radius=24, blur=30, alpha=80)
    od.rounded_rectangle(box, radius=24, fill=WHITE, outline=(220, 215, 205), width=2)
    od.text((cx0 + 30, cy0 + 30), "30-YEAR PROJECTION",
            font=f(34), fill=ACCENT, anchor="lt")
    img.alpha_composite(overlay)

    # Curve
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    lx0 = cx0 + 80
    ly0 = cy0 + 130
    lx1 = cx1 - 60
    ly1 = cy1 - 80
    pts = []
    n = 30
    for i in range(n):
        t = i / (n - 1)
        v = 0.05 + 0.95 * (1 - math.exp(-2.4 * t))
        pts.append((lx0 + t * (lx1 - lx0), ly1 - v * (ly1 - ly0)))
    poly = pts + [(lx1, ly1), (lx0, ly1)]
    od.polygon(poly, fill=(117, 230, 193, 130))
    od.line(pts, fill=PRIMARY, width=8)
    # Milestones
    for milestone, label in [(0.30, "$1K/mo"), (0.55, "$5K/mo"), (0.85, "FIRE")]:
        idx = int(milestone * (n - 1))
        px, py = pts[idx]
        od.ellipse((px - 18, py - 18, px + 18, py + 18), fill=HIGHLIGHT, outline=PRIMARY, width=4)
        od.text((px, py - 50), label, font=f(28), fill=PRIMARY, anchor="mm")
    img.alpha_composite(overlay)

    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
# Image 6: MOBILE PREVIEW
# ===========================================================================
def render_mobile(out: str) -> None:
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img, spacing=44, radius=4)
    d = ImageDraw.Draw(img)

    pill(img, SIZE // 2, 180, "WORKS EVERYWHERE",
         font=f(44), bg=PRIMARY, pad_x=60, pad_y=24)
    text_center(d, (SIZE // 2, 330), "EXCEL  ·  GOOGLE SHEETS", f(78), PRIMARY)
    text_center(d, (SIZE // 2, 420),
                "Open on desktop, edit on mobile, sync everywhere",
                f(34, bold=False), TEXT_MUTED)

    # Phone frame
    px, py = SIZE // 2, 1280
    pw, ph = 620, 1200
    phone_box = (px - pw // 2, py - ph // 2, px + pw // 2, py + ph // 2)
    drop_shadow(img, phone_box, radius=60, blur=50, alpha=120)
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    od.rounded_rectangle(phone_box, radius=60, fill=(25, 25, 30))
    # Inner screen
    bezel = 22
    screen = (phone_box[0] + bezel, phone_box[1] + bezel + 30,
              phone_box[2] - bezel, phone_box[3] - bezel - 30)
    od.rounded_rectangle(screen, radius=42, fill=BG)
    # Notch
    od.rounded_rectangle((px - 90, phone_box[1] + 16, px + 90, phone_box[1] + 48),
                         radius=18, fill=(15, 15, 18))

    # Phone app top bar
    sx0, sy0, sx1, sy1 = screen
    od.rounded_rectangle((sx0, sy0, sx1, sy0 + 90), radius=42, fill=PRIMARY)
    od.rectangle((sx0, sy0 + 50, sx1, sy0 + 90), fill=PRIMARY)
    od.text(((sx0 + sx1) // 2, sy0 + 50), "Dividend Wealth",
            font=f(32), fill=WHITE, anchor="mm")

    # Phone KPI cards (stacked)
    y = sy0 + 130
    for lab, val in [("PORTFOLIO VALUE", "$148,420"),
                     ("MONTHLY INCOME", "$612.40"),
                     ("YIELD %", "4.95%"),
                     ("YEARS TO FIRE", "14.8")]:
        cbox = (sx0 + 30, y, sx1 - 30, y + 140)
        od.rounded_rectangle(cbox, radius=18, fill=WHITE, outline=(220, 215, 205), width=2)
        od.text((cbox[0] + 26, y + 30), lab, font=f(22), fill=ACCENT, anchor="lt")
        od.text((cbox[0] + 26, y + 96), val, font=f(48), fill=PRIMARY, anchor="lm")
        y += 160

    # Mini bar chart
    bx0, by0 = sx0 + 50, y + 40
    bx1, by1 = sx1 - 50, y + 250
    od.text((bx0, by0 - 30), "MONTHLY INCOME", font=f(22), fill=ACCENT, anchor="lt")
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"]
    vals = [80, 110, 95, 140, 120, 160]
    bw = (bx1 - bx0) / (len(vals) * 1.6)
    max_v = max(vals)
    for i, v in enumerate(vals):
        x = bx0 + i * (bw * 1.6)
        h = (v / max_v) * (by1 - by0)
        od.rounded_rectangle((x, by1 - h, x + bw, by1), radius=4, fill=PRIMARY)
        od.text((x + bw / 2, by1 + 18), months[i],
                font=f(18), fill=TEXT_MUTED, anchor="mt")

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
        ("01_hero.png",        render_hero),
        ("02_dashboard.png",   render_dashboard_closeup),
        ("03_holdings.png",    render_holdings),
        ("04_charts.png",      render_charts),
        ("05_projections.png", render_projections),
        ("06_mobile.png",      render_mobile),
    ]
    for name, fn in targets:
        path = os.path.join(out_dir, name)
        fn(path)
        print(f"  ✓ {name}")
    print(f"Wrote {len(targets)} images to {out_dir}")


if __name__ == "__main__":
    main()

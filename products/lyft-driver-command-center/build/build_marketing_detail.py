"""Detailed marketing image set for Lyft Driver Command Center™ (4 images, 2000x2000).

Deeper, benefit-driven marketing beyond the app screenshots:

  07_features.png   - feature spotlights (what it does + why it matters)
  08_compare.png    - "basic mileage app" vs Lyft Driver Command Center™
  09_howitworks.png - up & running in 4 steps
  10_value.png      - what's included, who it's for & the guarantee

Reuses the branded helpers from build_marketing.py.
Run: python3 build_marketing_detail.py
"""
from __future__ import annotations
import os
import math
from PIL import Image, ImageDraw

from build_marketing import (
    SIZE, PRIMARY, PRIMARY_DK, PRIMARY_LT, ACCENT, GOLD, GOLD_LT, GOLD_HI, SURFACE,
    HIGHLIGHT, BG, WHITE, TEXT, TEXT_MUTED, DANGER, MINT_BG, WARN_BG, RED_BG, GRID,
    ROW_ALT, fs, fserif, premium_bg, pill, wordmark, gold_divider, wheel_crest,
    shadow, grad_round, radial_glow, tc, donut, legend, hbars, fit_font, vgradient,
)


def wrap(d, text, font, max_w):
    words = text.split(); lines = []; cur = ""
    for w in words:
        t = (cur + " " + w).strip()
        if d.textlength(t, font=font) <= max_w:
            cur = t
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


def card(img, box, radius=22, blur=16, dy=10, outline=(232, 224, 208)):
    shadow(img, box, radius, blur, 44, dy)
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
    od.rounded_rectangle(box, radius=radius, fill=WHITE, outline=outline, width=2)
    od.rounded_rectangle((box[0], box[1] + 8, box[0] + 7, box[3] - 8), radius=3, fill=GOLD_LT)
    img.alpha_composite(ov)


def mini_donut(img, cx, cy, r, pct, label, sub):
    d = ImageDraw.Draw(img)
    donut(d, cx, cy, r, [(pct, PRIMARY), (100 - pct, SURFACE)], label, sub)


def header_band(img, tagpill, title, subtitle, band_h=430):
    premium_bg(img, band_h=band_h)
    d = ImageDraw.Draw(img)
    wheel_crest(img, SIZE // 2, 110, r=48)
    pill(img, SIZE // 2, 214, tagpill, font=fs(24), pad_x=40, pad_y=18)
    tc(d, (SIZE // 2, 300), title, fserif(56), WHITE)
    gold_divider(img, SIZE // 2, 360, width=480)
    if subtitle:
        tc(d, (SIZE // 2, 400), subtitle, fs(23, bold=False), (226, 214, 190))


# ===========================================================================
def render_features(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    header_band(img, "BUILT TO GROW YOUR TAKE-HOME", "Where Drivers Win",
                "Four engines that run the business behind the wheel")
    feats = [
        ("$", "Take-Home Engine", "donut", (66, "66%", "kept"),
         "Every fare, tip and bonus minus fuel and expenses rolls into one honest net number — your real $21.68 per hour."),
        ("%", "Mileage & Taxes", "table", None,
         "Log business miles and the Tax Center auto-picks mileage vs actual and banks your quarterly set-aside — the deduction most drivers under-claim."),
        ("box", "Earnings Analytics", "stock", None,
         "Best shifts, best zones and best hours — see exactly where each dollar-per-hour comes from, then repeat it."),
        ("star", "Driver Health", "bars", None,
         "One 0-100% score blends net earnings, $/hour, trips, savings, consistency & your tax reserve."),
    ]
    margin = 84; gap = 28
    cw = (SIZE - 2 * margin - gap) // 2
    ch = 560
    top = 500
    for i, (badge, title, kind, dat, body) in enumerate(feats):
        r, c = divmod(i, 2)
        x = margin + c * (cw + gap); y = top + r * (ch + gap)
        card(img, (x, y, x + cw, y + ch))
        ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
        bx, by = x + 66, y + 66
        od.ellipse((bx - 40, by - 40, bx + 40, by + 40), fill=PRIMARY, outline=GOLD_HI, width=3)
        if kind == "stock":
            od.rounded_rectangle((bx - 22, by - 14, bx + 22, by + 20), radius=4, outline=GOLD_HI, width=4)
            od.line((bx - 22, by, bx + 22, by), fill=GOLD_HI, width=3)
            od.line((bx, by - 14, bx, by), fill=GOLD_HI, width=3)
        elif kind == "bars":
            pts = []
            for k in range(10):
                a = -math.pi / 2 + k * math.pi / 5
                rr = 26 if k % 2 == 0 else 11
                pts.append((bx + math.cos(a) * rr, by + math.sin(a) * rr))
            od.polygon(pts, fill=GOLD_HI)
        else:
            od.text((bx, by - 3), badge, font=fserif(40), fill=GOLD_HI, anchor="mm")
        od.text((x + 124, by - 16), title, font=fs(30), fill=PRIMARY, anchor="lm")
        od.text((x + 124, by + 22), "AUTOMATIC", font=fs(15), fill=ACCENT, anchor="lm")
        for li, line in enumerate(wrap(od, body, fs(20, bold=False), cw - 80)):
            od.text((x + 40, y + 148 + li * 32), line, font=fs(20, bold=False), fill=TEXT, anchor="lt")
        img.alpha_composite(ov)
        vy = y + ch - 180
        if kind == "donut":
            mini_donut(img, x + cw // 2, vy + 30, 96, dat[0], dat[1], dat[2])
        elif kind == "table":
            od = ImageDraw.Draw(img)
            rows = [("Mileage deduction", "$2,010"), ("Set aside (25%)", "$554"), ("Best method", "Mileage")]
            tx0 = x + 44; tx1 = x + cw - 44; ty = vy - 30
            grad_round(img, (tx0, ty, tx1, ty + 34), 6, PRIMARY_LT, PRIMARY_DK)
            for h, fx in zip(["TAX ITEM", "AMOUNT"], [0.0, 0.78]):
                od.text((tx0 + 14 + (tx1 - tx0) * fx, ty + 17), h, font=fs(15), fill=WHITE, anchor="lm")
            for ri, (a, b) in enumerate(rows):
                ry = ty + 34 + ri * 42
                if ri % 2:
                    od.rectangle((tx0, ry, tx1, ry + 42), fill=ROW_ALT)
                od.text((tx0 + 14, ry + 21), a, font=fs(17), fill=PRIMARY, anchor="lm")
                od.text((tx0 + 14 + (tx1 - tx0) * 0.78, ry + 21), b, font=fs(17), fill=PRIMARY, anchor="lm")
        elif kind == "stock":
            od = ImageDraw.Draw(img)
            items = [("Concert surge night", "$340", MINT_BG, PRIMARY), ("Downtown / 6th St", "$41/hr", MINT_BG, PRIMARY),
                     ("Airport (AUS)", "$34/hr", SURFACE, (110, 88, 58)), ("Weekend rush", "$310", MINT_BG, PRIMARY)]
            tx0 = x + 44; tx1 = x + cw - 44; ty = vy - 30
            for ri, (nm, st, bg, fg) in enumerate(items):
                ry = ty + ri * 44
                od.rounded_rectangle((tx0, ry, tx1, ry + 38), radius=8, fill=(246, 241, 232))
                od.text((tx0 + 16, ry + 19), nm, font=fs(18, bold=False), fill=TEXT, anchor="lm")
                od.rounded_rectangle((tx1 - 128, ry + 6, tx1 - 12, ry + 32), radius=13, fill=bg)
                od.text((tx1 - 70, ry + 19), st, font=fs(14), fill=fg, anchor="mm")
        elif kind == "bars":
            hbars(img, ImageDraw.Draw(img), (x + 44, y + 300, x + cw - 44, y + ch - 30),
                  [("Net vs goal", 0.93, "93%"), ("Net $/hour", 1.0, "100%"), ("Savings", 0.70, "70%"), ("Tax reserve", 0.78, "78%")])
    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
def render_compare(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    header_band(img, "WHY IT'S DIFFERENT", "More Than a Mileage App",
                "Everything a basic tracker leaves out — built in")
    rows = [
        ("Track miles & trips", True, True),
        ("True net $/hour & $/mile", False, True),
        ("Fuel & fill-up log", False, True),
        ("Vehicle & maintenance log", False, True),
        ("Business expenses + cost/mile", False, True),
        ("Mileage vs actual tax method", False, True),
        ("Quarterly tax set-aside", False, True),
        ("Monthly household budget", False, True),
        ("Savings buckets & goals", False, True),
        ("Bonuses, quests & hot zones", False, True),
        ("One Driver Health Score", False, True),
    ]
    margin = 110; top = 500
    tw = SIZE - 2 * margin
    c1 = margin + tw * 0.60; c2 = margin + tw * 0.83
    d = ImageDraw.Draw(img)
    hh = 74
    grad_round(img, (margin, top, SIZE - margin, top + hh), 12, PRIMARY_LT, PRIMARY_DK)
    d.text((margin + 24, top + hh // 2), "FEATURE", font=fs(22), fill=WHITE, anchor="lm")
    d.text((c1, top + hh // 2), "Basic\nMileage App", font=fs(18), fill=(210, 210, 205), anchor="mm")
    d.text((c2, top + hh // 2), "Command\nCenter™", font=fs(18), fill=GOLD_HI, anchor="mm")
    rh = 108
    for i, (label, a, b) in enumerate(rows):
        ry = top + hh + i * rh
        if i % 2:
            d.rounded_rectangle((margin, ry, SIZE - margin, ry + rh), radius=8, fill=ROW_ALT)
        d.text((margin + 24, ry + rh // 2), label, font=fs(23, bold=False), fill=TEXT, anchor="lm")
        if a:
            d.ellipse((c1 - 22, ry + rh // 2 - 22, c1 + 22, ry + rh // 2 + 22), outline=(180, 175, 165), width=3)
            d.line((c1 - 10, ry + rh // 2, c1 - 2, ry + rh // 2 + 10), fill=(150, 145, 135), width=4)
            d.line((c1 - 2, ry + rh // 2 + 10, c1 + 12, ry + rh // 2 - 8), fill=(150, 145, 135), width=4)
        else:
            d.ellipse((c1 - 22, ry + rh // 2 - 22, c1 + 22, ry + rh // 2 + 22), fill=RED_BG)
            d.line((c1 - 9, ry + rh // 2 - 9, c1 + 9, ry + rh // 2 + 9), fill=DANGER, width=4)
            d.line((c1 - 9, ry + rh // 2 + 9, c1 + 9, ry + rh // 2 - 9), fill=DANGER, width=4)
        d.ellipse((c2 - 24, ry + rh // 2 - 24, c2 + 24, ry + rh // 2 + 24), fill=PRIMARY)
        d.line((c2 - 11, ry + rh // 2 + 1, c2 - 3, ry + rh // 2 + 11), fill=HIGHLIGHT, width=5)
        d.line((c2 - 3, ry + rh // 2 + 11, c2 + 13, ry + rh // 2 - 10), fill=HIGHLIGHT, width=5)
    fy0 = top; fy1 = top + hh + len(rows) * rh
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
    od.rounded_rectangle((c2 - 90, fy0 - 8, c2 + 90, fy1 + 8), radius=16, outline=GOLD_LT, width=4)
    img.alpha_composite(ov)
    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
def render_howitworks(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    header_band(img, "SET UP IN MINUTES", "Up & Running in 4 Steps",
                "No spreadsheet skills needed — if you can drive, you can run it")
    steps = [
        ("1", "Download & open", "Buy once and instantly download your Excel file, plus a 1-click link to copy the Google Sheets version."),
        ("2", "Set your car & goals", "In Settings, add your vehicle, the IRS mileage rate and monthly goals. One time only."),
        ("3", "Log every shift", "After each shift, add hours, trips, fares, tips, miles & fuel in the Shift Log. It takes 30 seconds."),
        ("4", "Watch your net grow", "Net $/hour, the mileage deduction, your budget and Driver Health Score all update automatically."),
    ]
    margin = 100; top = 500
    sh = 300; gap = 24
    for i, (num, title, body) in enumerate(steps):
        y = top + i * (sh + gap)
        card(img, (margin, y, SIZE - margin, y + sh))
        ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
        cx, cy = margin + 130, y + sh // 2
        od.ellipse((cx - 70, cy - 70, cx + 70, cy + 70), fill=PRIMARY, outline=GOLD_HI, width=4)
        od.text((cx, cy - 4), num, font=fserif(72), fill=GOLD_HI, anchor="mm")
        tx = margin + 270
        od.text((tx, cy - 66), title, font=fs(34), fill=PRIMARY, anchor="lt")
        for li, line in enumerate(wrap(od, body, fs(22, bold=False), SIZE - margin - tx - 50)):
            od.text((tx, cy - 14 + li * 34), line, font=fs(22, bold=False), fill=TEXT, anchor="lt")
        img.alpha_composite(ov)
        if i < len(steps) - 1:
            ImageDraw.Draw(img).line((margin + 130, y + sh + 2, margin + 130, y + sh + gap - 2), fill=GOLD_LT, width=4)
    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
def render_value(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    header_band(img, "EVERYTHING YOU GET", "One File, Your Whole Business",
                "Instant download · lifetime access · free updates")
    d = ImageDraw.Draw(img)
    margin = 84
    lx0, lx1 = margin, SIZE // 2 - 14
    card(img, (lx0, 500, lx1, 1300))
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
    od.text((lx0 + 40, 540), "WHAT'S INCLUDED", font=fs(26), fill=ACCENT, anchor="lt")
    includes = [
        "18-tab Excel workbook (.xlsx)", "Google Sheets edition (1-click copy)",
        "12-KPI driver dashboard", "Shift log + earnings engine",
        "Mileage & fuel trackers", "Business expenses + Tax Center",
        "Monthly household budget", "Savings, goals & bonuses",
        "Start-Here quick-start guide (PDF)", "Free lifetime updates",
    ]
    for i, item in enumerate(includes):
        yy = 600 + i * 64
        od.ellipse((lx0 + 44, yy, lx0 + 76, yy + 32), fill=PRIMARY)
        od.line((lx0 + 52, yy + 16, lx0 + 59, yy + 24), fill=HIGHLIGHT, width=4)
        od.line((lx0 + 59, yy + 24, lx0 + 70, yy + 9), fill=HIGHLIGHT, width=4)
        od.text((lx0 + 96, yy + 16), item, font=fs(22, bold=False), fill=TEXT, anchor="lm")
    img.alpha_composite(ov)
    rx0, rx1 = SIZE // 2 + 14, SIZE - margin
    card(img, (rx0, 500, rx1, 890))
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
    od.text((rx0 + 40, 540), "WHO IT'S FOR", font=fs(26), fill=ACCENT, anchor="lt")
    who = ["Full-time rideshare drivers", "Part-time & weekend drivers", "Lyft & Uber drivers",
           "Delivery drivers (DoorDash, Uber Eats)", "Gig-economy side hustlers", "Anyone who deducts mileage"]
    for i, w in enumerate(who):
        yy = 604 + i * 44
        od.text((rx0 + 44, yy), "•", font=fs(24), fill=GOLD_LT, anchor="lm")
        od.text((rx0 + 72, yy), w, font=fs(21, bold=False), fill=TEXT, anchor="lm")
    img.alpha_composite(ov)
    card(img, (rx0, 916, rx1, 1300))
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
    od.text((rx0 + 40, 956), "WORKS WITH", font=fs(26), fill=ACCENT, anchor="lt")
    for i, t in enumerate(["✓ Microsoft Excel (2019 / 2021 / 365)", "✓ Google Sheets (free account)",
                            "✓ Windows, Mac, tablet & mobile"]):
        od.text((rx0 + 44, 1020 + i * 44), t, font=fs(21, bold=False), fill=TEXT, anchor="lm")
    od.rounded_rectangle((rx0 + 40, 1160, rx1 - 40, 1264), radius=14, fill=MINT_BG)
    od.text((rx0 + 64, 1188), "★  Not working right?", font=fs(21), fill=PRIMARY, anchor="lm")
    od.text((rx0 + 64, 1228), "Message me — I help within 24 hours.", font=fs(19, bold=False), fill=TEXT, anchor="lm")
    img.alpha_composite(ov)
    pill(img, SIZE // 2, 1420, "REPLACES A BOOKKEEPER, A TAX PREP & A BUDGET APP", font=fs(28), pad_x=44, pad_y=22,
         grad=(PRIMARY_LT, PRIMARY_DK))
    d = ImageDraw.Draw(img)
    tc(d, (SIZE // 2, 1530), "One point of control for your earnings AND the business behind the wheel.",
       fs(24, bold=False), TEXT_MUTED)
    wheel_crest(img, SIZE // 2, 1700, r=54)
    wordmark(img, SIZE // 2, 1850, "LYFT DRIVER COMMAND CENTER", 54, max_w=1540)
    img.convert("RGB").save(out, "PNG", optimize=True)


def main():
    out_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "marketing")
    os.makedirs(out_dir, exist_ok=True)
    targets = [
        ("07_features.png", render_features),
        ("08_compare.png", render_compare),
        ("09_howitworks.png", render_howitworks),
        ("10_value.png", render_value),
    ]
    for name, fn in targets:
        fn(os.path.join(out_dir, name))
        print(f"  ✓ {name}")
    print(f"Wrote {len(targets)} detailed images to {out_dir}")


if __name__ == "__main__":
    main()

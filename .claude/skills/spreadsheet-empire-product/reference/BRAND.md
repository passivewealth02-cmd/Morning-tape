# BRAND — the one visual identity every product shares

Premium, calm, "expensive software," never cutesy. Forest green + warm gold on
ivory. Serif wordmarks, sans body. It should look like a $200 SaaS dashboard, not
a free printable.

## Palette

| Token | Hex | RGB | Use |
| ----- | --- | --- | --- |
| PRIMARY | `#1B4F48` | 27,79,72 | headers, titles, primary fills |
| PRIMARY_DK / _LT | `#123833` / `#215C53` | 18,56,51 / 33,92,83 | gradients |
| ACCENT (Gold) | `#937356` | 147,115,86 | labels, accents |
| GOLD_LT / GOLD_HI | `#C9A86A` / `E0C48C` | 201,168,106 / 224,196,140 | rules, crest, chips |
| SURFACE | `#E5D3BA` | 229,211,186 | banners, totals, muted fills |
| HIGHLIGHT (Mint) | `#75E6C1` | 117,230,193 | positive bars/segments |
| MINT_BG / WARN_BG / RED_BG | `#E3F8EF` / `#FBF0E2` / `#FBE6E6` | status cell fills |
| IVORY / BG | `#FBF8F2` / 251,248,242 | page background |
| TEXT / TEXT_MUTED | `#333333` / 132,126,116 | body / captions |

Marketing page background is a subtle vertical ivory gradient with a faint dot
grid; the hero has a green top band with a gold rule (`premium_bg` + `hero_band`).

## Fonts (DejaVu — installed at /usr/share/fonts/truetype/dejavu/)

- `DejaVuSans-Bold` / `DejaVuSans` → all UI/body text (`fs()`).
- `DejaVuSerif-Bold` → wordmarks, big serif titles, donut centers (`fserif()`).

**Glyph safety** — DejaVu renders: `$ % ★(U+2605) ✓(U+2713) • ▲ ► —`. It does
**NOT** render color emoji (💰🎒🚗 etc.) or `■`/`⭐(U+2B50)`. In marketing images
use only safe glyphs or **vector-drawn icons**. (Excel/Sheets cells CAN show
emoji, so emoji in `luxe_header` titles inside the .xlsx are fine — just not in
Pillow images.)

## The crest (one unique emblem per product)

Same frame every time, unique icon inside. Recipe:

```python
def <thing>_crest(c, cx, cy, r=56, glow=True):
    if glow: radial_glow(c, cx, cy, int(r*2.1), GOLD_HI, 90)
    # optional 2nd glow in a niche accent colour (Lyft pink, Eats green, apple red)
    grad_round(c, (cx-r,cy-r,cx+r,cy+r), 22, PRIMARY_LT, PRIMARY_DK, outline=GOLD_LT, width=4)
    ov=Image.new("RGBA",c.size,(0,0,0,0)); d=ImageDraw.Draw(ov)
    d.rounded_rectangle((cx-r+10,cy-r+10,cx+r-10,cy+r-10), radius=16, outline=GOLD_HI, width=2)
    c.alpha_composite(ov)
    ov=Image.new("RGBA",c.size,(0,0,0,0)); d=ImageDraw.Draw(ov)
    # ---- draw the niche icon in GOLD_HI here (polygons / arcs / ellipses) ----
    # add ONE small niche-accent colour touch (a pink hub, a green chevron, a red apple)
    c.alpha_composite(ov)
```

Crests shipped so far (don't reuse — invent a new one): play-button (YouTube),
globe (vacation), compass (road trip), chef toque (restaurant), music-note+glitch
(TikTok), camera-gradient (Instagram), microphone+waves (podcast), steering wheel
+pink hub (Lyft), delivery bag+green chevron (Uber Eats), carrot (Instacart),
backpack+green pocket (back-to-school), open book+red apple (homeschool).

## Layout constants (marketing)

- Canvas `SIZE = 2000` (square, Etsy).
- Hero: crest ~y132, tagline pill ~y256, `wordmark` ~y400, `gold_divider` ~y500,
  subtitle ~y550, 3 `stat_chip`s ~y704, `app_window(70,800,1930,1900)`, bottom
  `pill` ~y1948.
- `app_window` draws a mac-style window with a green sidebar listing `TABS`
  (needs module-level `TABS` + `FILE_LABEL`) and calls your `content_fn` for the
  right pane.
- Sheet-spread images: green band ~360 tall, pill+serif title+subtitle, then
  `app_window(70,400,1930,1930)` with the active tab highlighted.

## Print (PDF) style

Ink-light: white page, one green header band (h≈340) with a gold rule, gold
section bars (`SURFACE`), thin gray rules for fill-in lines, green table headers.
US Letter @ 300 dpi = 2550×3300. Footer: family/persona left, brand right.

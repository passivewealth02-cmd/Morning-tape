# Maintena — Brand Kit

The operational coordination layer for property maintenance. Feel: **fast, modern,
AI-assisted, operationally focused, effortless.** Inspiration: Linear, Stripe, Notion.

---

## 1. Logo

**The mark** is a rounded **M** monogram (a single rounded stroke) — usually reversed in
white on an indigo rounded tile.

Files (in `public/brand/`):
| File | Use |
|---|---|
| `maintena-icon.svg` | App icon / avatar — indigo tile + white M (rounded square) |
| `maintena-mark-white.svg` | The M alone, white — for dark/indigo backgrounds |
| `maintena-mark-indigo.svg` | The M alone, indigo — for white backgrounds |
| `maintena-lockup-light.svg` | Tile + "Maintena" wordmark, dark text — for light backgrounds |
| `maintena-lockup-dark.svg` | Tile + "Maintena" wordmark, white text — for dark backgrounds/video |
| `twitter-avatar.png`, `twitter-banner.png` | Social profile assets |

**Clear space:** keep at least the height of the tile's corner radius (≈15% of tile size)
clear on all sides. **Minimum size:** 24 px tall for the tile; 20 px for the standalone mark.

**Do:** use the white mark on indigo or dark; use the indigo tile on white/light.
**Don't:** recolor the mark outside the brand indigo/white, add gradients or shadows, stretch
it, rotate it, or place the tile on a busy photo without a solid backing.

---

## 2. Color

**Primary — Indigo**
| Token | Hex | Use |
|---|---|---|
| Indigo 600 (Primary) | `#4F46E5` | Logo tile, primary buttons, brand accents |
| Indigo 700 | `#4338CA` | Hover / pressed |
| Indigo 500 | `#6366F1` | Highlights, gradients |
| Indigo 100 | `#E0E7FF` | Soft fills |
| Indigo 50 | `#EEF2FF` | Section backgrounds, badges |

**Neutrals — Gray/Ink**
| Token | Hex | Use |
|---|---|---|
| Gray 900 (Ink) | `#111827` | Headlines |
| Gray 600 | `#4B5563` | Body text |
| Gray 500 | `#6B7280` | Secondary text |
| Gray 400 | `#9CA3AF` | Muted / captions |
| Gray 200 | `#E5E7EB` | Borders |
| Gray 100 | `#F3F4F6` | Dividers |
| Gray 50 | `#F9FAFB` | Surfaces / section backgrounds |
| White | `#FFFFFF` | Cards, reversed text |

**Status accents** (use sparingly, functional only)
| Meaning | Hex |
|---|---|
| Emergency / error | `#DC2626` |
| High / warning | `#EA580C` |
| Caution / trial | `#D97706` |
| Success / done | `#16A34A` |

---

## 3. Typography

**Primary typeface: Inter** (clean, modern, operational). Fallback: system sans-serif.
- Headlines: **600 (Semibold)**, tight tracking (`letter-spacing: -0.02em`)
- Body: 400–500, normal tracking
- Labels/eyebrows: 600, uppercase, wide tracking, small

Scale (web): H1 48–60px · H2 30px · H3 18px · Body 16–18px · Small 13–14px.

---

## 4. Voice & tone

Operational, plain-spoken, benefit-led. Talk about outcomes, not features.

- ✅ "Stop losing maintenance requests and chasing vendors."
- ✅ "Tenants scan. AI dispatches. You approve."
- ❌ "AI property management software" (too generic — the PRD flags this)
- ❌ Corporate filler ("leverage synergies", "robust solutions")

Tagline options: **"The AI operations layer for property maintenance."** ·
"Scan. Dispatch. Done." · "Maintenance coordination, on autopilot."

---

## 5. Remotion assets

### Color + font constants
```ts
// brand.ts
export const BRAND = {
  indigo: '#4F46E5',
  indigoDark: '#4338CA',
  indigoLight: '#6366F1',
  indigo50: '#EEF2FF',
  ink: '#111827',
  gray600: '#4B5563',
  gray400: '#9CA3AF',
  gray50: '#F9FAFB',
  white: '#FFFFFF',
  font: 'Inter, sans-serif',
}
```
Load Inter in Remotion with `@remotion/google-fonts/Inter`.

### Animated logo component (draws the M, then reveals the wordmark)
```tsx
import { interpolate, spring, useCurrentFrame, useVideoConfig } from 'remotion'

export const MaintenaLogo: React.FC<{ wordmark?: boolean }> = ({ wordmark = true }) => {
  const frame = useCurrentFrame()
  const { fps } = useVideoConfig()

  // Draw the M stroke (dash offset 0 -> full over ~20 frames)
  const draw = interpolate(frame, [0, 20], [1, 0], { extrapolateRight: 'clamp' })
  // Tile pops in with a spring
  const pop = spring({ frame, fps, config: { damping: 14 } })
  // Wordmark fades/slides in after the mark
  const word = interpolate(frame, [18, 32], [0, 1], { extrapolateLeft: 'clamp', extrapolateRight: 'clamp' })

  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 24, transform: `scale(${pop})` }}>
      <div style={{ width: 160, height: 160, borderRadius: 36, background: '#4F46E5',
                    display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
        <svg viewBox="0 0 120 120" width="80" height="80" fill="none">
          <path d="M 24 96 L 24 30 L 60 70 L 96 30 L 96 96" stroke="#fff" strokeWidth={16}
                strokeLinecap="round" strokeLinejoin="round"
                pathLength={1} strokeDasharray={1} strokeDashoffset={draw} />
        </svg>
      </div>
      {wordmark && (
        <span style={{ fontFamily: 'Inter, sans-serif', fontWeight: 600, fontSize: 88,
                       letterSpacing: -3, color: '#111827', opacity: word,
                       transform: `translateX(${(1 - word) * 20}px)` }}>
          Maintena
        </span>
      )}
    </div>
  )
}
```

### Motion guidelines
- **Easing:** spring or ease-out; nothing linear. Snappy, confident, not bouncy-cartoonish.
- **Backgrounds:** white `#FFFFFF` or ink `#111827`; use indigo `#4F46E5` for one hero beat.
- **Accent motion:** indigo underlines/pills that wipe in left-to-right.
- **Pace:** fast cuts (2–3s per beat). Operational, not sleepy.

### 20-second promo storyboard (script)
1. **0–3s** — Black screen, text types: *"Maintenance requests are chaos."* (calls/texts/emails icons scatter)
2. **3–6s** — They collapse into one clean ticket. VO/caption: *"Maintena fixes that."*
3. **6–9s** — Phone scans a QR on a door → form auto-fills unit. Caption: *"Tenants scan to report."*
4. **9–12s** — AI label snaps on: *Plumbing · Emergency · Plumber.* Caption: *"AI triages in seconds."*
5. **12–15s** — Vendor card slides in, "Assigned ✓". Caption: *"The right vendor, dispatched."*
6. **15–18s** — Tracking stepper fills Received → Done. Caption: *"Everyone stays in the loop."*
7. **18–20s** — Logo draw-in (component above) on white, tagline: **"The AI operations layer for property maintenance."** + `trymaintena.com`.

Formats to render: 1920×1080 (landscape), 1080×1920 (reels/stories), 1080×1080 (square).

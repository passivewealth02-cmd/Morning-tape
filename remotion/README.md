# Maintena — Promo Video (Remotion)

A 20-second brand promo built with [Remotion](https://remotion.dev). Renders in three
aspect ratios from the same timeline.

## Setup

```bash
cd remotion
npm install
```

## Preview (live studio)

```bash
npm run dev
```

Opens Remotion Studio in your browser — scrub the timeline, tweak any scene in `src/scenes/`,
and see changes live.

## Render to MP4

```bash
npm run render            # 1920x1080 landscape  -> out/maintena-landscape.mp4
npm run render:vertical   # 1080x1920 reels/story -> out/maintena-vertical.mp4
npm run render:square     # 1080x1080 square      -> out/maintena-square.mp4
```

## Structure

| File | What it is |
|---|---|
| `src/index.ts` | Registers the root |
| `src/Root.tsx` | The three compositions (landscape / vertical / square) |
| `src/Promo.tsx` | The 20s timeline — 7 scenes stitched with `<Series>` |
| `src/brand.ts` | Brand colors + Inter font (single source of truth) |
| `src/MaintenaLogo.tsx` | Animated logo (draws the M, reveals the wordmark) |
| `src/ui.tsx` | Shared `Caption` and `Chip` helpers |
| `src/scenes/*` | One file per beat |

## The 7 beats (30fps)

1. **Chaos** (0–3s) — "Maintenance requests are chaos" over scattered channels
2. **One ticket** (3–6s) — it all collapses into a single clean ticket
3. **Scan** (6–9s) — phone scans a QR, the form auto-fills the unit
4. **Triage** (9–12s) — AI labels snap on (Plumbing · Emergency · Plumber)
5. **Dispatch** (12–15s) — vendor card slides in, "Assigned ✓"
6. **Tracking** (15–18s) — the status stepper fills to Completed
7. **Logo** (18–20s) — animated logo + tagline + trymaintena.com

## Tweaking

- **Colors/fonts:** edit `src/brand.ts`
- **Copy:** each scene's text lives in its own file under `src/scenes/`
- **Timing:** change `durationInFrames` on each `<Series.Sequence>` in `src/Promo.tsx`
- **Length/fps/size:** edit `src/Root.tsx`
- **Voiceover/music:** drop an audio file in `public/` and add `<Audio src={staticFile('...')} />` to `Promo.tsx`

This folder is self-contained and excluded from the Next.js app build — it never affects the site.

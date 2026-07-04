# Design Autopilot Coach

A private, beginner-friendly design coach for Etsy PNG sellers. It turns a
rough idea (a phrase + a few tags) into a complete, opinionated **design
recipe**: layout, object placement, font pairing, colour palette with roles,
white-space notes, a design score, warnings, and a copy-paste AI prompt.

Live at **`/coach`**.

## What ships in this MVP

Everything runs **client-side and deterministically** — no API key, no
network, no cost. The advice comes from a hand-curated rules engine, so the
same input always produces the same recipe.

Implemented against the PRD:

- **Design Idea form** (Screen 2) — phrase, niche, product, style, objects,
  colour vibe, season, buyer, notes, with per-niche object suggestions.
- **Design Recipe output** (Screen 3 / PRD §13 template) — summary, layout,
  canvas, text placement, object placement map, font pairing, colour recipe,
  white-space notes, warnings, AI prompt, and Canva/Kittl build steps.
- **Layout Recipe Generator** (Feature 2) — 7 layouts, chosen from vibe +
  product + object count.
- **Object Placement Engine** (Feature 3) — roles (main / supporting /
  filler / frame / background), placements, sizes and rotations.
- **Font Pairing Assistant** (Feature 4) — 2-font pairings per vibe with
  hierarchy notes and warnings.
- **Colour Coach** (Feature 5) — named palettes with real hex swatches and a
  role + "where it goes" for every colour.
- **Design Score** (Feature 8) — before/after breakdown across 6 axes.
- **Auto-Fix suggestions** (Feature 9) — surfaced as warnings + exact steps.
- **Generate 3 Better Versions** (Feature 11) — Clean / Trendy / Playful.
- **AI Prompt Generator** (Feature 12).
- **Generate image** — renders the recipe into an actual transparent-background
  PNG via an image model (OpenAI `gpt-image-1` by default). Optional: set
  `OPENAI_API_KEY` to enable it. Without a key the button shows a friendly
  "copy the prompt into your own tool" message, so the coach still works.
- Saved recipes via `localStorage`.

## Reference images

The left rail has a **Reference images** uploader (up to 10). Images are
downscaled client-side (max 1024px, JPEG) so the payload stays small, kept as
data URLs in component state, and passed to the image route. When present, the
route uses the image-**edit** endpoint (`/images/edits`) so the model treats
them as a visual brief for palette, layout and style — otherwise it does a
plain text-to-image generation. References are a UI/generation input only;
they don't change the deterministic text recipe.

## Image generation

`lib/design-coach/image.ts` calls an OpenAI-style `/images/generations`
endpoint (or `/images/edits` when references are supplied) and returns a
base64 PNG data URL. `app/api/coach/image/route.ts`
exposes it as `POST /api/coach/image { prompt }`. Configure with:

| Env | Default | Purpose |
| --- | --- | --- |
| `OPENAI_API_KEY` | — | Required to enable image generation. |
| `COACH_IMAGE_MODEL` | `gpt-image-1` | Swap the model. |
| `COACH_IMAGE_BASE_URL` | OpenAI API | Point at any OpenAI-compatible host. |

The route returns `422 {code:"no_key"}` when unconfigured, `502` on provider
errors, and `{dataUrl, model}` on success. `gpt-image-1` is chosen because it
supports `background: "transparent"` — exactly what print-on-demand sellers
need.

## Code map

| Path | Role |
| --- | --- |
| `lib/design-coach/types.ts` | All shared types. |
| `lib/design-coach/knowledge.ts` | The curated catalog: vibes, layouts, palettes, fonts, object roles, niche hints. |
| `lib/design-coach/engine.ts` | Pure `generateRecipe(input)` rules engine + scoring, warnings, prompt, variants. |
| `lib/design-coach/index.ts` | Barrel + example presets + niche→object suggestions. |
| `app/coach/page.tsx` | Route + metadata. |
| `components/coach/coach-app.tsx` | The interactive form + recipe UI. |

## Extending it

- **New vibe/palette/fonts:** add an entry to `VIBES`, `PALETTES`,
  `FONT_PAIRINGS`, and `LAYOUT_PREFERENCE` in `knowledge.ts`.
- **New objects:** add to `OBJECT_CATALOG` with a role + rotation + size.
- **New niche:** add to `NICHE_HINTS` (keywords, suggested objects, default
  vibe).

## Not in this MVP (PRD "future")

The editable canvas editor with one-click physical Auto-Fix (Feature 10,
Screen 7) and image-upload critique (Screen 4, needs vision) are intentionally
out of scope for the text-based MVP. An optional AI-enhanced prompt path could
later layer on top of `lib/anthropic.ts`, but is not required — the engine is
complete on its own.

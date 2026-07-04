// The Design Autopilot Coach rules engine.
// Pure, deterministic functions that turn a beginner's DesignInput into a
// complete DesignRecipe. No network, no API key — this is the MVP brain.

import type {
  DesignInput,
  DesignRecipe,
  LayoutRecipe,
  LayoutId,
  ObjectRole,
  PlacedObject,
  TextHierarchy,
  ColourRecipe,
  FontPairing,
  DesignScore,
  ScoreBreakdown,
  DesignVariant,
} from './types'
import {
  VIBES,
  type VibeId,
  LAYOUTS,
  LAYOUT_PREFERENCE,
  PALETTES,
  FONT_PAIRINGS,
  OBJECT_CATALOG,
  ROLE_DEFAULTS,
  NICHE_HINTS,
  STOP_WORDS,
} from './knowledge'

// ---------------------------------------------------------------------------
// Normalization
// ---------------------------------------------------------------------------

function norm(s: string): string {
  return s.toLowerCase().trim()
}

/** Best-guess the vibe from the style text, then colour vibe, then niche. */
export function resolveVibe(input: DesignInput): VibeId {
  const haystacks = [input.style, input.colourVibe, input.notes].map(norm)

  for (const [vibe, def] of Object.entries(VIBES) as [VibeId, { keywords: string[] }][]) {
    for (const kw of def.keywords) {
      if (haystacks.some(h => h.includes(kw))) return vibe
    }
  }

  // Fall back to the niche's default vibe.
  const niche = norm(input.niche)
  const hint = NICHE_HINTS.find(h => h.keywords.some(k => niche.includes(k)))
  if (hint) return hint.defaultVibe

  return 'cute-retro'
}

function matchNiche(input: DesignInput) {
  const niche = norm(input.niche)
  const phrase = norm(input.phrase)
  return NICHE_HINTS.find(h => h.keywords.some(k => niche.includes(k) || phrase.includes(k)))
}

// ---------------------------------------------------------------------------
// Layout selection
// ---------------------------------------------------------------------------

export function pickLayout(input: DesignInput, vibe: VibeId): LayoutRecipe {
  const prefs = LAYOUT_PREFERENCE[vibe]
  const objectCount = input.objects.filter(Boolean).length
  const product = norm(input.productType)
  const style = norm(input.style)

  let id: LayoutId = prefs[0]

  // Nudges based on the design's shape.
  if (objectCount >= 4 && prefs.includes('three-frame')) {
    id = 'three-frame'
  } else if (/badge|circle|crest|logo/.test(style)) {
    id = 'badge-circle'
  } else if (/arch|rainbow|groovy/.test(style) && prefs.includes('retro-arch')) {
    id = 'retro-arch'
  } else if (/sticker|mug|tumbler/.test(product) && objectCount <= 2 && prefs.includes('center-object')) {
    id = 'center-object'
  }

  return { id, ...LAYOUTS[id] }
}

// ---------------------------------------------------------------------------
// Text hierarchy — pick the hero word.
// ---------------------------------------------------------------------------

export function buildTextHierarchy(input: DesignInput): TextHierarchy {
  const raw = input.phrase.trim()
  const words = raw.split(/\s+/).filter(Boolean)

  if (words.length === 0) {
    return {
      heroWord: 'YOUR WORD',
      leadText: '',
      tailText: '',
      optionalSubText: '',
      notes: 'Add a phrase and the coach will pick the strongest word to feature.',
    }
  }

  if (words.length === 1) {
    return {
      heroWord: words[0].toUpperCase(),
      leadText: '',
      tailText: '',
      optionalSubText: 'Add a tiny tagline below, like the niche or occasion.',
      notes: 'Single-word designs are strong — make that word large and bold.',
    }
  }

  // The last content word is almost always the punchline / hero in Etsy
  // quote designs ("Teaching is my JAM", "But first COFFEE"). Prefer it,
  // and only fall back to the longest content word if the last word is
  // very short filler (2 letters or fewer).
  const candidates = words
    .map((w, i) => ({ w, i, clean: w.replace(/[^a-z']/gi, '') }))
    .filter(c => c.clean && !STOP_WORDS.has(c.clean.toLowerCase()))

  let heroIndex = words.length - 1
  if (candidates.length > 0) {
    const lastContent = candidates[candidates.length - 1]
    if (lastContent.clean.length <= 2) {
      heroIndex = candidates.reduce((a, b) => (b.clean.length > a.clean.length ? b : a)).i
    } else {
      heroIndex = lastContent.i
    }
  }

  const heroWord = words[heroIndex].replace(/[^a-z'0-9]/gi, '').toUpperCase()
  const leadText = words.slice(0, heroIndex).join(' ')
  const tailText = words.slice(heroIndex + 1).join(' ')

  return {
    heroWord,
    leadText,
    tailText,
    optionalSubText: input.season || input.niche ? `Optional tiny line: "${(input.season || input.niche).toLowerCase()}"` : '',
    notes: `"${heroWord}" is the punch word — make it the biggest, boldest element. The rest of the phrase stays smaller and supports it.`,
  }
}

// ---------------------------------------------------------------------------
// Object placement engine
// ---------------------------------------------------------------------------

function lookupObject(name: string) {
  const key = norm(name)
  if (OBJECT_CATALOG[key]) return OBJECT_CATALOG[key]
  // loose contains match (e.g. "red apple" -> apple, "tiny stars" -> stars)
  const hit = Object.keys(OBJECT_CATALOG).find(k => key.includes(k))
  return hit ? OBJECT_CATALOG[hit] : null
}

const PLACEMENT_BY_LAYOUT: Record<LayoutId, Partial<Record<ObjectRole, string[]>>> = {
  'stacked-typography': {
    main: ['tucked slightly behind and below the hero word'],
    supporting: ['on the lower-left, angled inward', 'on the lower-right, angled inward'],
    filler: ['a tiny cluster in the top-right corner', 'a tiny cluster in the bottom-left corner'],
    frame: ['along the lower-left edge', 'along the lower-right edge'],
    background: ['faded behind the whole text block'],
  },
  'center-object': {
    main: ['dead-center, the largest element'],
    supporting: ['resting at the base of the main object, left', 'resting at the base of the main object, right'],
    filler: ['tiny accents in two opposite corners'],
    frame: ['wrapping the lower edge of the object'],
    background: ['a soft arc behind the object'],
  },
  'badge-circle': {
    main: ['centered inside the badge'],
    supporting: ['at the 9 o\'clock inner edge', 'at the 3 o\'clock inner edge'],
    filler: ['small dots where the top and bottom text meet'],
    frame: ['forming the badge ring'],
    background: ['a faint fill inside the circle'],
  },
  'retro-arch': {
    main: ['centered directly under the arch'],
    supporting: ['lower-left beneath the object', 'lower-right beneath the object'],
    filler: ['a few sparkles in the negative space under the arch'],
    frame: ['tracing the underside of the arch'],
    background: ['a sunburst behind the arch'],
  },
  'three-frame': {
    main: ['integrated with the center text block'],
    supporting: ['centered in the left third', 'centered in the right third'],
    filler: ['small accents in the outer corners'],
    frame: ['bordering the left and right zones'],
    background: ['a faded band behind the center zone'],
  },
  'wreath-corner': {
    main: ['inside the wreath opening with the text'],
    supporting: ['denser cluster at the top-left corner', 'denser cluster at the bottom-right corner'],
    filler: ['small buds trailing off the clusters'],
    frame: ['forming the wreath around the words'],
    background: ['a pale wash inside the opening'],
  },
  'minimal-typography': {
    main: ['one tiny accent beside the hero word, or omit it'],
    supporting: ['omit — restraint is the style'],
    filler: ['omit'],
    frame: ['a single thin underline, optional'],
    background: ['none'],
  },
}

export function placeObjects(input: DesignInput, layout: LayoutRecipe): PlacedObject[] {
  const objects = input.objects.map(o => o.trim()).filter(Boolean)
  if (objects.length === 0) return []

  // Determine each object's role, then make sure exactly one "main" wins.
  type Draft = { name: string; role: ObjectRole; rotation: string; size: string }
  const drafts: Draft[] = objects.map(name => {
    const meta = lookupObject(name) ?? ROLE_DEFAULTS.supporting
    return { name, role: meta.role, rotation: meta.rotation, size: meta.size }
  })

  // Minimal layout suppresses supporting/filler objects.
  const isMinimal = layout.id === 'minimal-typography'

  // Exactly one main: keep the first "main", demote the rest to supporting.
  let mainSeen = false
  for (const d of drafts) {
    if (d.role === 'main') {
      if (mainSeen) {
        d.role = 'supporting'
        d.rotation = ROLE_DEFAULTS.supporting.rotation
        d.size = ROLE_DEFAULTS.supporting.size
      } else {
        mainSeen = true
      }
    }
  }
  // If nothing is a main, promote the first non-filler object.
  if (!mainSeen) {
    const promote = drafts.find(d => d.role !== 'filler' && d.role !== 'background') ?? drafts[0]
    promote.role = 'main'
    promote.rotation = ROLE_DEFAULTS.main.rotation
    promote.size = ROLE_DEFAULTS.main.size
  }

  // Assign concrete placements, cycling through the slots available per role.
  const counters: Record<string, number> = {}
  return drafts.map(d => {
    const role: ObjectRole = isMinimal && d.role !== 'main' ? 'filler' : d.role
    const slots = PLACEMENT_BY_LAYOUT[layout.id][role] ?? ['in an empty area away from the text']
    const idx = (counters[role] = (counters[role] ?? 0)) % slots.length
    counters[role] = (counters[role] ?? 0) + 1
    return {
      name: d.name,
      role: d.role,
      placement: slots[idx],
      size: d.size,
      rotation: d.rotation,
    }
  })
}

// ---------------------------------------------------------------------------
// White space + warnings
// ---------------------------------------------------------------------------

export function whiteSpaceNotes(input: DesignInput, objects: PlacedObject[]): string[] {
  const notes = [
    'Leave clear breathing room around the hero word — nothing should touch the letters.',
    'Keep the outer 10% margin clean; pull every element inward off the edges.',
    'Balance the visual weight left-to-right so one side doesn\'t feel heavier.',
  ]
  const fillerCount = objects.filter(o => o.role === 'filler').length
  if (fillerCount >= 3) {
    notes.push('You have a lot of filler — keep it to small clusters in the top-right and bottom-left only, and remove 30–40% of it.')
  }
  if (objects.length >= 5) {
    notes.push('With this many elements, cut anything that isn\'t earning its place. One clear focal point beats a busy canvas.')
  }
  if (objects.length <= 1) {
    notes.push('The design is sparse — that\'s fine for a minimal look, but make the hero word big enough to fill the space confidently.')
  }
  return notes
}

export function buildWarnings(input: DesignInput, objects: PlacedObject[], vibe: VibeId): string[] {
  const warnings: string[] = []
  const mainObj = objects.find(o => o.role === 'main')

  if (mainObj) {
    warnings.push(`Do not make the ${mainObj.name} larger than the hero word — the object supports the text, it doesn't compete with it.`)
  }
  const supporting = objects.filter(o => o.role === 'supporting')
  if (supporting.length > 0) {
    warnings.push('Keep supporting-object angles between 10 and 25 degrees. Avoid random extreme rotations.')
  }
  if (objects.filter(o => o.role === 'filler').length >= 4) {
    warnings.push('Too many filler shapes will crowd the letters — thin them out and keep them tiny.')
  }
  // Colour readability nudge tied to the palette.
  if (vibe === 'cozy' || vibe === 'cute-retro' || vibe === 'boho') {
    warnings.push('Don\'t set light text on cream without a dark outline — add a chocolate/dark-anchor outline so it stays readable.')
  }
  if (vibe === 'kids') {
    warnings.push('Bright colours only read as "clean" with a bold dark outline on every element — don\'t skip the outlines.')
  }
  warnings.push('Check the hero word is still readable shrunk to Etsy thumbnail size — if it isn\'t, make it bigger or bolder.')
  if (input.objects.filter(Boolean).length === 0) {
    warnings.push('No objects listed — that works for a typographic design, but consider one small accent to give the eye a focal point.')
  }
  return warnings
}

// ---------------------------------------------------------------------------
// Scoring (Feature 8) — a heuristic "typical beginner attempt" score vs.
// the "after following this recipe" score.
// ---------------------------------------------------------------------------

function clamp(n: number, lo: number, hi: number) {
  return Math.max(lo, Math.min(hi, n))
}

export function scoreDesign(input: DesignInput, objects: PlacedObject[]): DesignScore {
  const objectCount = objects.length
  const fillerCount = objects.filter(o => o.role === 'filler').length
  const wordCount = input.phrase.trim().split(/\s+/).filter(Boolean).length

  // "Before" models common beginner mistakes given the inputs.
  const before: ScoreBreakdown = {
    overall: 0,
    readability: clamp(8 - Math.floor(fillerCount / 2) - (wordCount > 5 ? 2 : 0), 3, 8),
    colourHarmony: clamp(6 - (objectCount > 4 ? 1 : 0), 3, 7),
    spacing: clamp(8 - fillerCount - (objectCount > 4 ? 1 : 0), 3, 8),
    fontPairing: 5, // beginners tend to over-use fonts
    objectBalance: clamp(8 - Math.abs(objectCount - 3), 3, 8),
    thumbnailReadability: clamp(8 - (wordCount > 4 ? 2 : 0) - Math.floor(fillerCount / 2), 3, 8),
  }
  before.overall = Math.round(
    ((before.readability + before.colourHarmony + before.spacing + before.fontPairing + before.objectBalance + before.thumbnailReadability) / 60) * 100,
  )

  // "After" assumes the recipe's rules are applied — strong, consistent.
  const after: ScoreBreakdown = {
    overall: 0,
    readability: 9,
    colourHarmony: 9,
    spacing: 9,
    fontPairing: 9,
    objectBalance: clamp(9 - (objectCount > 6 ? 1 : 0), 7, 9),
    thumbnailReadability: 9,
  }
  after.overall = Math.round(
    ((after.readability + after.colourHarmony + after.spacing + after.fontPairing + after.objectBalance + after.thumbnailReadability) / 60) * 100,
  )

  const mainIssues: string[] = []
  if (before.fontPairing <= 5) mainIssues.push('Likely too many fonts — cut down to two.')
  if (before.spacing <= 6) mainIssues.push('Elements are probably crowding the hero word — add breathing room.')
  if (before.colourHarmony <= 6) mainIssues.push('Colours need clearer roles and a dark anchor for contrast.')
  if (before.thumbnailReadability <= 6) mainIssues.push('The hero word may not read at thumbnail size — make it larger.')
  if (fillerCount >= 3) mainIssues.push('Filler shapes are too close to the letters — thin them out.')
  if (mainIssues.length === 0) mainIssues.push('Solid start — the recipe tightens spacing, colour roles and hierarchy.')

  return { before, after, mainIssues }
}

// ---------------------------------------------------------------------------
// AI prompt (Feature 12) + Canva build steps.
// ---------------------------------------------------------------------------

export function buildAiPrompt(
  input: DesignInput,
  layout: LayoutRecipe,
  hero: TextHierarchy,
  objects: PlacedObject[],
  palette: ColourRecipe,
  fonts: FontPairing,
): string {
  const styleLabel = norm(input.style) || 'clean'
  const parts: string[] = []

  parts.push(
    `Create a ${styleLabel} ${input.niche || ''} PNG design for a ${input.productType || 'shirt'} with the phrase "${input.phrase}".`.replace(/\s+/g, ' '),
  )
  parts.push(`Use a ${layout.name.toLowerCase()} with "${hero.heroWord}" as the largest ${layout.id === 'retro-arch' ? 'arched' : 'center'} word${hero.leadText ? `, and "${hero.leadText}" smaller ${layout.id === 'retro-arch' ? 'beneath' : 'above'} it` : ''}.`)

  if (objects.length) {
    const objText = objects
      .map(o => `${o.name} ${o.placement}${o.rotation && o.rotation !== 'none' && o.rotation !== 'none (upright)' ? ` (rotate ${o.rotation})` : ''}`)
      .join(', ')
    parts.push(`Add ${objText}.`)
  }

  const anchor = palette.swatches.find(s => /anchor/i.test(s.role))
  const main = palette.swatches.find(s => /main text/i.test(s.role))
  const sec = palette.swatches.find(s => /secondary/i.test(s.role))
  const accent = palette.swatches.find(s => /accent/i.test(s.role))
  parts.push(
    `Use the "${palette.name}" palette: ${main?.name.toLowerCase()} for the main word, ${anchor?.name.toLowerCase()} outlines, ${sec?.name.toLowerCase()} highlights${accent ? `, and ${accent.name.toLowerCase()} for tiny accents only` : ''}.`,
  )
  parts.push(`Use a ${fonts.mainFont.style.toLowerCase()} for the hero word and a ${fonts.secondaryFont.style.toLowerCase()} for the smaller text. Two fonts maximum.`)
  parts.push('Keep the design balanced with clean white space around the main phrase. Transparent background. No product mockup. No watermark.')

  return parts.join(' ')
}

export function buildCanvaSteps(
  input: DesignInput,
  layout: LayoutRecipe,
  hero: TextHierarchy,
  objects: PlacedObject[],
  canvasSize: string,
): string[] {
  const steps = [`Create a ${canvasSize} transparent canvas.`]
  steps.push(`Add "${hero.heroWord}" first, centered — this is your biggest, boldest word.`)
  if (hero.leadText) steps.push(`Add "${hero.leadText}" smaller, ${layout.id === 'retro-arch' ? 'straight under the arch' : 'above the hero word'}.`)
  if (hero.tailText) steps.push(`Add "${hero.tailText}" as a small line below the hero word.`)

  const main = objects.find(o => o.role === 'main')
  if (main) steps.push(`Add the ${main.name} ${main.placement}, sized ${main.size}.`)
  objects.filter(o => o.role === 'supporting').forEach(o => steps.push(`Add the ${o.name} ${o.placement}${o.rotation !== 'none' ? `, rotated to ${o.rotation}` : ''}.`))
  const fillers = objects.filter(o => o.role === 'filler')
  if (fillers.length) steps.push(`Add ${fillers.map(f => f.name).join(', ')} only in the empty corners — keep them tiny.`)

  steps.push('Apply the colour palette, assigning each colour to its role.')
  steps.push('Shrink the canvas preview to thumbnail size and confirm the hero word still reads clearly.')
  steps.push('Export as a transparent PNG.')
  return steps
}

// ---------------------------------------------------------------------------
// Three better versions (Feature 11).
// ---------------------------------------------------------------------------

function variantFor(
  id: DesignVariant['id'],
  title: string,
  tagline: string,
  input: DesignInput,
  layoutId: LayoutId,
  vibe: VibeId,
): DesignVariant {
  const layout = { id: layoutId, ...LAYOUTS[layoutId] }
  const hero = buildTextHierarchy(input)
  const objects = placeObjects(input, layout)
  const palette = PALETTES[vibe]
  const fonts = FONT_PAIRINGS[vibe]
  return {
    id,
    title,
    tagline,
    layout: layout.name,
    textPlacement: `"${hero.heroWord}" leads; ${layout.mainTextArea.toLowerCase()}`,
    objectPlacement: objects.length
      ? objects.slice(0, 3).map(o => `${o.name} ${o.placement}`).join('; ')
      : 'Typographic — no objects, let the words carry it.',
    fontPairing: `${fonts.name} (${fonts.mainFont.style} + ${fonts.secondaryFont.style})`,
    palette: `${palette.name} — ${palette.swatches.map(s => s.name).slice(0, 4).join(', ')}`,
    warning: buildWarnings(input, objects, vibe)[0] ?? 'Keep one clear focal point.',
    aiPrompt: buildAiPrompt(input, layout, hero, objects, palette, fonts),
  }
}

export function buildVariants(input: DesignInput, primaryVibe: VibeId, primaryLayout: LayoutId): DesignVariant[] {
  // A: clean beginner-safe minimal-ish layout.
  const cleanLayout: LayoutId = primaryLayout === 'stacked-typography' ? 'minimal-typography' : 'stacked-typography'
  // C: playful/unique alt layout.
  const playfulLayout: LayoutId =
    primaryLayout === 'center-object' ? 'three-frame' : primaryLayout === 'retro-arch' ? 'center-object' : 'retro-arch'

  return [
    variantFor('clean', 'Clean Beginner Layout', 'Safest to execute — hardest to mess up.', input, cleanLayout, primaryVibe),
    variantFor('trendy', 'Trendy Etsy Layout', 'What\'s selling on Etsy right now.', input, primaryLayout, primaryVibe),
    variantFor('playful', 'Unique Playful Layout', 'A little more personality to stand out.', input, playfulLayout, primaryVibe),
  ]
}

// ---------------------------------------------------------------------------
// Top-level: assemble the full recipe.
// ---------------------------------------------------------------------------

function canvasFor(productType: string): { size: string; safeMargin: string; background: string } {
  const p = norm(productType)
  if (/sticker/.test(p)) return { size: '3000 x 3000 px', safeMargin: '10% safe margin', background: 'transparent' }
  if (/mug|tumbler/.test(p)) return { size: '2550 x 1050 px (wrap)', safeMargin: '10% safe margin, keep art off the seam', background: 'transparent' }
  if (/tote|bag/.test(p)) return { size: '3600 x 3600 px', safeMargin: '12% safe margin', background: 'transparent' }
  // default shirt/PNG
  return { size: '4500 x 5400 px', safeMargin: '10% safe margin', background: 'transparent' }
}

function summarize(input: DesignInput, layout: LayoutRecipe, vibe: VibeId): string {
  const vibeLabel = VIBES[vibe].label.toLowerCase()
  const objs = input.objects.filter(Boolean)
  const objText = objs.length ? ` with ${objs.slice(0, 3).join(', ')} framing the words` : ''
  return `A ${vibeLabel} ${input.niche || 'design'} for a ${input.productType || 'PNG'} built on a ${layout.name.toLowerCase()}${objText}. The hero word leads, everything else supports it.`
}

export function generateRecipe(input: DesignInput): DesignRecipe {
  const vibe = resolveVibe(input)
  const layout = pickLayout(input, vibe)
  const hero = buildTextHierarchy(input)
  const objects = placeObjects(input, layout)
  const palette = PALETTES[vibe]
  const fonts = FONT_PAIRINGS[vibe]
  const canvas = canvasFor(input.productType)

  return {
    input,
    summary: summarize(input, layout, vibe),
    layout,
    canvas,
    textHierarchy: hero,
    objects,
    fonts,
    colours: palette,
    whiteSpaceNotes: whiteSpaceNotes(input, objects),
    warnings: buildWarnings(input, objects, vibe),
    score: scoreDesign(input, objects),
    aiPrompt: buildAiPrompt(input, layout, hero, objects, palette, fonts),
    buildSteps: buildCanvaSteps(input, layout, hero, objects, canvas.size),
    variants: buildVariants(input, vibe, layout.id),
  }
}

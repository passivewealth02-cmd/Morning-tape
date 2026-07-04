// The design knowledge base for Design Autopilot Coach.
// This is a hand-curated rules catalog — layouts, colour palettes, font
// pairings and object roles — that the engine composes into a recipe.
// Everything is keyed by a small set of normalized "vibes" and "niches"
// so a beginner's free-text answers still map onto solid design advice.

import type { LayoutId, ObjectRole, LayoutRecipe, FontPairing, ColourRecipe } from './types'

// ---------------------------------------------------------------------------
// Vibes — the normalized style buckets everything else keys off of.
// ---------------------------------------------------------------------------

export type VibeId =
  | 'cute-retro'
  | 'groovy-retro'
  | 'western'
  | 'soft-feminine'
  | 'kids'
  | 'minimal'
  | 'boho'
  | 'cozy'
  | 'varsity'
  | 'coquette'
  | 'sketch'
  | 'glam'
  | 'edgy'
  | 'cartoon'
  | 'bright-doodle'
  | 'watercolor'

export const VIBES: Record<VibeId, { label: string; keywords: string[] }> = {
  'cute-retro': { label: 'Cute Retro', keywords: ['cute retro', 'cute', 'retro cute', 'playful retro', 'warm retro', 'retro'] },
  'groovy-retro': { label: 'Groovy / Trendy Retro', keywords: ['groovy', 'trendy', 'wavy', '70s', 'seventies', 'hippie', 'psychedelic', 'retro groovy'] },
  western: { label: 'Western', keywords: ['western', 'cowboy', 'cowgirl', 'rodeo', 'desert', 'country', 'howdy'] },
  'soft-feminine': { label: 'Soft Feminine', keywords: ['feminine', 'soft', 'floral', 'pretty', 'elegant', 'romantic', 'wedding', 'girly'] },
  kids: { label: 'Kids / Bubbly', keywords: ['kids', 'kid', 'baby', 'bubble', 'bubbly', 'fun', 'birthday'] },
  minimal: { label: 'Minimal / Modern', keywords: ['minimal', 'modern', 'clean', 'simple', 'neutral', 'classy', 'chic', 'aesthetic'] },
  boho: { label: 'Boho / Earthy', keywords: ['boho', 'earthy', 'natural', 'sage', 'muted', 'organic', 'terracotta'] },
  cozy: { label: 'Cozy / Vintage', keywords: ['cozy', 'cosy', 'vintage', 'warm', 'coffee', 'fall', 'autumn', 'cabin', 'christmas'] },
  varsity: { label: 'Varsity / Collegiate', keywords: ['varsity', 'collegiate', 'athletic', 'college', 'sporty', 'block', 'arched'] },
  coquette: { label: 'Coquette / Line Art', keywords: ['coquette', 'line art', 'lineart', 'dainty', 'ribbon', 'delicate', 'single color', 'minimal line', 'sketch line'] },
  sketch: { label: 'Sketch / Marker Doodle', keywords: ['sketch', 'doodle', 'marker', 'hand drawn', 'handdrawn', 'scribble', 'two tone', 'two-tone', 'brush'] },
  glam: { label: 'Leopard Glam', keywords: ['leopard', 'glam', 'cheetah', 'animal print', 'glitter', 'luxe', 'glitzy'] },
  edgy: { label: 'Edgy / Checkered', keywords: ['edgy', 'checkered', 'checkerboard', 'grunge', 'punk', 'graffiti', 'y2k'] },
  cartoon: { label: 'Cartoon Characters', keywords: ['cartoon', 'character', 'characters', 'mascot', 'kawaii', 'silly', 'cute food', 'smiley'] },
  'bright-doodle': { label: 'Bright Doodle Letters', keywords: ['colorful', 'colourful', 'bright', 'multicolor', 'multicolour', 'puffy', 'preppy', 'hand lettered', 'rainbow letters'] },
  watercolor: { label: 'Watercolor Floral', keywords: ['watercolor', 'watercolour', 'painted', 'painterly', 'botanical', 'soft floral', 'personalized'] },
}

// ---------------------------------------------------------------------------
// Layouts (PRD Feature 2).
// ---------------------------------------------------------------------------

export const LAYOUTS: Record<LayoutId, Omit<LayoutRecipe, 'id'>> = {
  'stacked-typography': {
    name: 'Stacked Typography Layout',
    why: 'Text-forward designs (shirts, quotes, teacher/mom/nurse/coffee) read best when the words are the artwork and objects only support them.',
    mainTextArea: 'Center of the canvas — the hero word sits dead-center as the single focal point.',
    smallTextArea: 'A smaller line directly above the hero word, optionally curved.',
    mainObjectArea: 'Tucked slightly behind or beneath the hero word so the text stays on top.',
    supportingObjectArea: 'One on the lower-left, one on the lower-right, framing the base of the text.',
    fillerArea: 'Only the top-right and bottom-left empty corners.',
    keepEmpty: 'The outer 10% margin and the space immediately hugging each letter.',
  },
  'center-object': {
    name: 'Center Object Layout',
    why: 'When a single cute character or object is the star (coffee cup, animal, mug pun), it belongs in the middle with text wrapped around it.',
    mainTextArea: 'A short line above the object and/or a short line below it.',
    smallTextArea: 'Below the main object as a tagline.',
    mainObjectArea: 'Dead-center, the largest element on the canvas.',
    supportingObjectArea: 'Small items resting at the object\'s base or peeking from behind it.',
    fillerArea: 'Tiny accents in the four corners, sparingly.',
    keepEmpty: 'Even breathing room on all four sides so the object feels centered.',
  },
  'badge-circle': {
    name: 'Badge / Circle Layout',
    why: 'Professional and school/nurse/sports niches look polished when contained in a circular badge — it reads as a logo or crest.',
    mainTextArea: 'Curved along the top inner edge of the circle.',
    smallTextArea: 'Curved along the bottom inner edge, mirroring the top.',
    mainObjectArea: 'Centered inside the circle.',
    supportingObjectArea: 'Small dots or stars marking the left and right where the top and bottom text meet.',
    fillerArea: 'A thin decorative ring or dashed border, no loose filler outside the badge.',
    keepEmpty: 'Everything outside the circle stays clean and empty.',
  },
  'retro-arch': {
    name: 'Retro Arch Layout',
    why: 'Trendy Etsy shirts sell with an arched/rainbow word — it instantly reads as modern and retro.',
    mainTextArea: 'Arched across the top like a rainbow, the largest element.',
    smallTextArea: 'Straight line centered beneath the arch.',
    mainObjectArea: 'Centered under the arch, framed by the curve of the text.',
    supportingObjectArea: 'Balanced on the lower-left and lower-right beneath the object.',
    fillerArea: 'A few tiny sparkles inside the negative space under the arch.',
    keepEmpty: 'The area directly under the peak of the arch, so it frames the object.',
  },
  'three-frame': {
    name: 'Three-Frame Element Layout',
    why: 'When there are several objects, giving each a clear zone around the text stops the design from turning into clutter.',
    mainTextArea: 'Centered, occupying the middle third.',
    smallTextArea: 'Directly above or below the hero word within the center zone.',
    mainObjectArea: 'Center zone, integrated with the text.',
    supportingObjectArea: 'One object in the left third, one in the right third, at equal size.',
    fillerArea: 'Small accents only where the three zones meet the corners.',
    keepEmpty: 'Clear gutters between the three zones so they don\'t collide.',
  },
  'wreath-corner': {
    name: 'Wreath / Corner Frame Layout',
    why: 'Floral, feminine, seasonal and cozy designs feel finished when greenery/flowers frame the words like a wreath.',
    mainTextArea: 'Centered inside the wreath opening.',
    smallTextArea: 'A second line inside the opening, above or below the hero word.',
    mainObjectArea: 'The wreath itself forms the frame around the text.',
    supportingObjectArea: 'Denser floral clusters at two opposite corners (top-left and bottom-right).',
    fillerArea: 'Small leaves and buds trailing off the main clusters.',
    keepEmpty: 'The center opening around the text stays clear so the words read.',
  },
  'minimal-typography': {
    name: 'Minimal Typography Layout',
    why: 'Modern, neutral and simple designs are strongest with almost no decoration — the type and the white space do the work.',
    mainTextArea: 'Centered, with generous space above and below.',
    smallTextArea: 'A single small line above or below, widely letter-spaced.',
    mainObjectArea: 'One tiny line-art accent, or none at all.',
    supportingObjectArea: 'None — restraint is the style.',
    fillerArea: 'None.',
    keepEmpty: 'The majority of the canvas — let the design breathe.',
  },
}

// Which layouts suit which vibe/niche, best-first. The engine also nudges
// based on object count.
export const LAYOUT_PREFERENCE: Record<VibeId, LayoutId[]> = {
  'cute-retro': ['stacked-typography', 'retro-arch', 'center-object'],
  'groovy-retro': ['retro-arch', 'stacked-typography', 'center-object'],
  western: ['badge-circle', 'stacked-typography', 'retro-arch'],
  'soft-feminine': ['wreath-corner', 'stacked-typography', 'minimal-typography'],
  kids: ['center-object', 'stacked-typography', 'three-frame'],
  minimal: ['minimal-typography', 'stacked-typography', 'badge-circle'],
  boho: ['wreath-corner', 'minimal-typography', 'stacked-typography'],
  cozy: ['stacked-typography', 'wreath-corner', 'center-object'],
  varsity: ['retro-arch', 'stacked-typography', 'minimal-typography'],
  coquette: ['center-object', 'three-frame', 'minimal-typography'],
  sketch: ['center-object', 'stacked-typography', 'three-frame'],
  glam: ['stacked-typography', 'retro-arch', 'center-object'],
  edgy: ['stacked-typography', 'retro-arch', 'three-frame'],
  cartoon: ['three-frame', 'center-object', 'stacked-typography'],
  'bright-doodle': ['stacked-typography', 'retro-arch', 'three-frame'],
  watercolor: ['center-object', 'wreath-corner', 'stacked-typography'],
}

// ---------------------------------------------------------------------------
// Colour palettes (PRD Feature 5). Real hex values so the UI can render
// swatches. Each swatch carries a role and where-to-use note.
// ---------------------------------------------------------------------------

export const PALETTES: Record<VibeId, ColourRecipe> = {
  'cute-retro': {
    name: 'Retro Classroom',
    swatches: [
      { role: 'Dark anchor', name: 'Chocolate brown', hex: '#4A2E1E', usage: 'Outlines and the smallest readable text.' },
      { role: 'Main text', name: 'Burnt orange', hex: '#C65D2E', usage: 'The largest hero word.' },
      { role: 'Secondary text', name: 'Cream', hex: '#F4E6C9', usage: 'Highlights and the supporting line of text.' },
      { role: 'Accent', name: 'Teal', hex: '#2C7A7B', usage: 'Tiny stars and small accents only.' },
      { role: 'Highlight', name: 'Mustard yellow', hex: '#E0A82E', usage: 'A drop shadow or secondary accent.' },
      { role: 'Object colour', name: 'Apple red', hex: '#B23A2E', usage: 'The main object (e.g. an apple).' },
    ],
    usageNotes: [
      'Use chocolate brown for every outline so the design holds together.',
      'Reserve teal for the tiniest accents — it should never compete with the hero word.',
      'Use cream instead of pure white; it reads warmer and more retro.',
    ],
  },
  'groovy-retro': {
    name: 'Groovy Sunset',
    swatches: [
      { role: 'Dark anchor', name: 'Deep aubergine', hex: '#3B2A3A', usage: 'Outlines and small text.' },
      { role: 'Main text', name: 'Retro orange', hex: '#E2703A', usage: 'The hero / arched word.' },
      { role: 'Secondary text', name: 'Butter cream', hex: '#F6E7C1', usage: 'Supporting line and highlights.' },
      { role: 'Accent', name: 'Avocado green', hex: '#6B8E23', usage: 'Wavy underlines and small shapes.' },
      { role: 'Highlight', name: 'Golden yellow', hex: '#EBB434', usage: 'Sunburst rays or shadow layer.' },
      { role: 'Object colour', name: 'Dusty pink', hex: '#D98A7E', usage: 'Groovy flowers or accent objects.' },
    ],
    usageNotes: [
      'Keep the arch word in retro orange with a chocolate/aubergine outline for contrast.',
      'Use avocado and golden yellow together sparingly — one leads, one supports.',
      'Wavy shapes and sunbursts read best in a single accent colour, not rainbow.',
    ],
  },
  western: {
    name: 'Desert Rodeo',
    swatches: [
      { role: 'Dark anchor', name: 'Espresso brown', hex: '#3E2C23', usage: 'Outlines and small text.' },
      { role: 'Main text', name: 'Rust red', hex: '#A8412A', usage: 'The hero word.' },
      { role: 'Secondary text', name: 'Sand', hex: '#E7D3AE', usage: 'Supporting line and badge fill.' },
      { role: 'Accent', name: 'Turquoise', hex: '#3AA6A0', usage: 'Small western accents (a star, a rope tie).' },
      { role: 'Highlight', name: 'Golden tan', hex: '#C89B4A', usage: 'Badge ring and shadow.' },
      { role: 'Object colour', name: 'Cactus green', hex: '#5E7C4E', usage: 'Cactus or desert objects.' },
    ],
    usageNotes: [
      'A badge ring in golden tan on a sand fill gives the instant western/crest look.',
      'Turquoise is the classic western pop — use it in one small spot only.',
      'Keep the palette warm and dusty; avoid bright, clean primaries.',
    ],
  },
  'soft-feminine': {
    name: 'Blush Garden',
    swatches: [
      { role: 'Dark anchor', name: 'Plum', hex: '#5A3A4A', usage: 'Outlines and small readable text.' },
      { role: 'Main text', name: 'Dusty rose', hex: '#C77F86', usage: 'The hero word.' },
      { role: 'Secondary text', name: 'Ivory', hex: '#F6EFE7', usage: 'Highlights and supporting text.' },
      { role: 'Accent', name: 'Sage green', hex: '#8BA382', usage: 'Leaves and greenery in the wreath.' },
      { role: 'Highlight', name: 'Soft gold', hex: '#CDA96A', usage: 'Delicate accents and thin shadows.' },
      { role: 'Object colour', name: 'Muted coral', hex: '#D98E7A', usage: 'Flower blooms.' },
    ],
    usageNotes: [
      'Let sage green greenery frame the words while blooms stay muted, not neon.',
      'Use plum for text outlines so pastel colours still read clearly.',
      'Keep it low-saturation throughout — softness is the whole appeal.',
    ],
  },
  kids: {
    name: 'Bubble Pop',
    swatches: [
      { role: 'Dark anchor', name: 'Navy', hex: '#243B6B', usage: 'Bold outlines and small text.' },
      { role: 'Main text', name: 'Cherry red', hex: '#E23B3B', usage: 'The hero word.' },
      { role: 'Secondary text', name: 'Sky blue', hex: '#5BB8E6', usage: 'Supporting text and shapes.' },
      { role: 'Accent', name: 'Sunny yellow', hex: '#F6C445', usage: 'Stars, dots and bursts.' },
      { role: 'Highlight', name: 'Grass green', hex: '#57B85A', usage: 'Ground line or accents.' },
      { role: 'Object colour', name: 'Bright orange', hex: '#F2812F', usage: 'The main character or object.' },
    ],
    usageNotes: [
      'Thick navy outlines make bright colours read as fun and clean rather than messy.',
      'Bright primaries are welcome here — just cap it at 3–4 so it stays balanced.',
      'Give every element a chunky outline for a sticker-ready look.',
    ],
  },
  minimal: {
    name: 'Mono Minimal',
    swatches: [
      { role: 'Dark anchor', name: 'Ink black', hex: '#1E1E1E', usage: 'The hero word and outlines.' },
      { role: 'Main text', name: 'Charcoal', hex: '#3A3A3A', usage: 'Hero word (soft alternative to pure black).' },
      { role: 'Secondary text', name: 'Warm grey', hex: '#8A857D', usage: 'Supporting text.' },
      { role: 'Accent', name: 'Terracotta', hex: '#B9694E', usage: 'One tiny accent — a dot or a single line.' },
      { role: 'Highlight', name: 'Bone', hex: '#EFEAE1', usage: 'A subtle background block if needed.' },
      { role: 'Object colour', name: 'Stone', hex: '#B7B0A5', usage: 'A single line-art accent, if any.' },
    ],
    usageNotes: [
      'One accent colour, used once. Everything else stays in the neutral range.',
      'Contrast comes from size and space, not colour — make the hero word big.',
      'Skip shadows and outlines; crisp edges suit the modern look.',
    ],
  },
  boho: {
    name: 'Earthy Boho',
    swatches: [
      { role: 'Dark anchor', name: 'Cocoa', hex: '#4B382A', usage: 'Outlines and small text.' },
      { role: 'Main text', name: 'Terracotta', hex: '#B5643C', usage: 'The hero word.' },
      { role: 'Secondary text', name: 'Oat', hex: '#EADDC7', usage: 'Supporting text and highlights.' },
      { role: 'Accent', name: 'Sage', hex: '#8B9A78', usage: 'Leaves and small botanicals.' },
      { role: 'Highlight', name: 'Ochre', hex: '#C79A46', usage: 'Sun shapes and thin accents.' },
      { role: 'Object colour', name: 'Clay pink', hex: '#C98E7C', usage: 'Florals and desert objects.' },
    ],
    usageNotes: [
      'Stay in warm, muted earth tones — no bright or cool colours.',
      'Sage and terracotta are the signature boho pair; let ochre support.',
      'Use cocoa outlines lightly, or skip outlines for a soft watercolour feel.',
    ],
  },
  cozy: {
    name: 'Cozy Cocoa',
    swatches: [
      { role: 'Dark anchor', name: 'Dark roast brown', hex: '#3A281E', usage: 'Outlines and small text.' },
      { role: 'Main text', name: 'Cinnamon', hex: '#A25A32', usage: 'The hero word.' },
      { role: 'Secondary text', name: 'Cream', hex: '#F1E4CC', usage: 'Highlights and supporting text.' },
      { role: 'Accent', name: 'Forest green', hex: '#3E5E45', usage: 'Pine, holly or small accents.' },
      { role: 'Highlight', name: 'Caramel', hex: '#CB914B', usage: 'Steam swirls and warm shadow.' },
      { role: 'Object colour', name: 'Berry red', hex: '#9E3B36', usage: 'Berries, mugs or seasonal objects.' },
    ],
    usageNotes: [
      'Cream over pure white keeps it warm and cozy rather than clinical.',
      'Forest green and berry red make it feel seasonal — dial them back for year-round niches.',
      'Layer a soft caramel shadow behind the hero word for warmth.',
    ],
  },
  varsity: {
    name: 'Varsity Pastel',
    swatches: [
      { role: 'Dark anchor', name: 'Slate navy', hex: '#33415C', usage: 'The block-letter outline and any tiny text.' },
      { role: 'Main text', name: 'Powder pink', hex: '#F3B6C6', usage: 'The big collegiate block letters.' },
      { role: 'Secondary text', name: 'Mint', hex: '#9FD6C9', usage: 'The connector word (e.g. the script “and”).' },
      { role: 'Accent', name: 'Sky blue', hex: '#8FC1E3', usage: 'A second outline pass or small accent.' },
      { role: 'Highlight', name: 'Cream', hex: '#FBF3E4', usage: 'Inner highlight and grain texture base.' },
      { role: 'Object colour', name: 'Soft coral', hex: '#E79A9A', usage: 'Any small object or underline.' },
    ],
    usageNotes: [
      'Give the block letters a double outline (mint over navy) — that’s the collegiate look.',
      'Add a light speckled/distressed texture inside the letters, not over the outline.',
      'Keep it to two soft tones plus the outline so it stays clean at thumbnail size.',
    ],
  },
  coquette: {
    name: 'Coquette Ribbon',
    swatches: [
      { role: 'Dark anchor', name: 'Denim blue', hex: '#3E5C8A', usage: 'The single line-art colour for icons and script.' },
      { role: 'Main text', name: 'Denim blue', hex: '#3E5C8A', usage: 'The script phrase, same delicate line weight.' },
      { role: 'Secondary text', name: 'Dusty blue', hex: '#7C9CC4', usage: 'A lighter second line if needed.' },
      { role: 'Accent', name: 'Blush pink', hex: '#E7B7C4', usage: 'One soft blush touch — a bow or heart.' },
      { role: 'Highlight', name: 'Cream', hex: '#F6EFE7', usage: 'Leave the background clean/transparent.' },
      { role: 'Object colour', name: 'Denim blue', hex: '#3E5C8A', usage: 'Line-art icons (bow, boot, hat, horseshoe).' },
    ],
    usageNotes: [
      'Stay one-colour and thin-lined — coquette is about delicate line art, not fills.',
      'A small row of line-art icons above a flowing script is the signature layout.',
      'Add at most one blush accent so it stays soft and minimal.',
    ],
  },
  sketch: {
    name: 'Two-Tone Marker',
    swatches: [
      { role: 'Dark anchor', name: 'Cobalt blue', hex: '#1E6FE0', usage: 'Hand-drawn object outlines and one text colour.' },
      { role: 'Main text', name: 'Bright orange', hex: '#F0501E', usage: 'The marker-brush hero words.' },
      { role: 'Secondary text', name: 'Cobalt blue', hex: '#1E6FE0', usage: 'The alternating words in the phrase.' },
      { role: 'Accent', name: 'Sky blue', hex: '#5FA8F0', usage: 'Small fills inside objects (e.g. coffee).' },
      { role: 'Highlight', name: 'Off white', hex: '#FBFAF6', usage: 'Keep the background clean/transparent.' },
      { role: 'Object colour', name: 'Cobalt blue', hex: '#1E6FE0', usage: 'Doodled book, mug, heart line art.' },
    ],
    usageNotes: [
      'Commit to exactly two marker colours — alternate them word by word.',
      'Everything looks hand-drawn: wobbly brush letters and loose doodle icons.',
      'Overlap the doodle objects slightly with the words for an energetic sketchbook feel.',
    ],
  },
  glam: {
    name: 'Leopard Glam',
    swatches: [
      { role: 'Dark anchor', name: 'Black', hex: '#1A1A1A', usage: 'Outlines, the script word, and leopard spots.' },
      { role: 'Main text', name: 'Blush pink', hex: '#E7A9B4', usage: 'One or two solid words.' },
      { role: 'Secondary text', name: 'Warm tan', hex: '#C9A57B', usage: 'The alternating solid word.' },
      { role: 'Accent', name: 'Leopard gold', hex: '#B98A4E', usage: 'The leopard-print fill base.' },
      { role: 'Highlight', name: 'Cream', hex: '#F3E9DA', usage: 'Inner glow behind the letters.' },
      { role: 'Object colour', name: 'Black', hex: '#1A1A1A', usage: 'Hearts, stars and the pencil underline.' },
    ],
    usageNotes: [
      'Put leopard print inside ONE word and keep the rest solid — too much print gets noisy.',
      'Outline everything in black so blush and tan stay crisp and readable.',
      'Neutrals + one blush pop is the glam formula; add tiny leopard hearts as accents.',
    ],
  },
  edgy: {
    name: 'Hot Pink Checker',
    swatches: [
      { role: 'Dark anchor', name: 'Ink black', hex: '#141414', usage: 'Grunge brush words, outlines, checkerboard squares.' },
      { role: 'Main text', name: 'Hot pink', hex: '#E6197F', usage: 'The loud brush/spray words.' },
      { role: 'Secondary text', name: 'Ink black', hex: '#141414', usage: 'The alternating grunge word.' },
      { role: 'Accent', name: 'White', hex: '#FFFFFF', usage: 'The other half of the checkerboard.' },
      { role: 'Highlight', name: 'Bubble pink', hex: '#F49FC6', usage: 'A soft glow or second pink.' },
      { role: 'Object colour', name: 'Hot pink', hex: '#E6197F', usage: 'Apple, pencil, lightning, crown, stars.' },
    ],
    usageNotes: [
      'Mix textures on purpose: checkerboard fill on some letters, spray/brush on others.',
      'Keep it to hot pink + black + white so the chaos still reads as one design.',
      'Doodle extras (crown, lightning bolt, scribble stars) sell the edgy/Y2K vibe.',
    ],
  },
  cartoon: {
    name: 'Cartoon Pop',
    swatches: [
      { role: 'Dark anchor', name: 'Ink black', hex: '#2B2B2B', usage: 'Thin character outlines, faces, arms and legs.' },
      { role: 'Main text', name: 'Leaf green', hex: '#6AA84F', usage: 'Any title text or a character.' },
      { role: 'Secondary text', name: 'Warm tan', hex: '#D9A46A', usage: 'Labels and secondary bits.' },
      { role: 'Accent', name: 'Sunny yellow', hex: '#F2C744', usage: 'Little stars and sparkles.' },
      { role: 'Highlight', name: 'Cheek pink', hex: '#F4A9B8', usage: 'Rosy cheeks and small hearts.' },
      { role: 'Object colour', name: 'Fresh green', hex: '#7BAE3F', usage: 'The main cute characters.' },
    ],
    usageNotes: [
      'The cute characters ARE the design — keep any text small and let them lead.',
      'Give every character the same thin black outline, dot eyes, a smile and rosy cheeks.',
      'Space characters evenly in a row with tiny stars filling the gaps.',
    ],
  },
  'bright-doodle': {
    name: 'Bright Doodle',
    swatches: [
      { role: 'Dark anchor', name: 'Gold', hex: '#D4A72C', usage: 'The glittery outline around every letter.' },
      { role: 'Main text', name: 'Hot pink', hex: '#EC5C8D', usage: 'Some of the puffy letters.' },
      { role: 'Secondary text', name: 'Periwinkle', hex: '#9B8FE4', usage: 'More of the letters (each one differs).' },
      { role: 'Accent', name: 'Tangerine', hex: '#F26B21', usage: 'More letters plus small accents.' },
      { role: 'Highlight', name: 'Butter yellow', hex: '#F6D65B', usage: 'More letters and glow.' },
      { role: 'Object colour', name: 'Grass green', hex: '#6FBF73', usage: 'Doodle hearts, flowers, smileys.' },
    ],
    usageNotes: [
      'Make every letter a different bright colour, all wrapped in the same gold glitter outline.',
      'Scatter tiny doodles — hearts, flowers, stars, a smiley — around the words.',
      'The gold outline is what ties the rainbow of letters together; keep it on all of them.',
    ],
  },
  watercolor: {
    name: 'Watercolor Garden',
    swatches: [
      { role: 'Dark anchor', name: 'Soft charcoal', hex: '#4A4A4A', usage: 'The thin script name and any small labels.' },
      { role: 'Main text', name: 'Rose pink', hex: '#E38AA0', usage: 'A warm floral tone for accents/text.' },
      { role: 'Secondary text', name: 'Sage green', hex: '#8FB07A', usage: 'Leaves and stems.' },
      { role: 'Accent', name: 'Lavender', hex: '#A98CD1', usage: 'Some blooms and small flowers.' },
      { role: 'Highlight', name: 'Buttercup', hex: '#F3C64B', usage: 'Sunny flower centres and warmth.' },
      { role: 'Object colour', name: 'Coral', hex: '#EE9B7A', usage: 'Painted books, petals and objects.' },
    ],
    usageNotes: [
      'Everything is soft painterly washes — no hard outlines, let colours bleed gently.',
      'Flowers and greenery grow up and out of the central object (books, a mug, a name).',
      'Set the name/phrase in a thin dark script so it stays readable over the soft colour.',
    ],
  },
}

// ---------------------------------------------------------------------------
// Font pairings (PRD Feature 4).
// ---------------------------------------------------------------------------

export const FONT_PAIRINGS: Record<VibeId, FontPairing> = {
  'cute-retro': {
    name: 'Chunky + Handwritten',
    mainFont: { role: 'Hero word', style: 'Chunky rounded retro display', example: 'Bogart, Cooper, or a bubbly retro serif' },
    secondaryFont: { role: 'Supporting line', style: 'Simple handwritten script', example: 'A clean marker or handwriting font' },
    warnings: ['Keep it to 2 fonts.', 'Never set the hero word in the script — it must be the boldest, most readable font.'],
    hierarchyNote: 'Hero word is largest and boldest; the handwritten line stays small and secondary.',
  },
  'groovy-retro': {
    name: 'Groovy Bold + Clean Sans',
    mainFont: { role: 'Hero / arch word', style: 'Groovy bold display (tall, wavy, 70s)', example: 'A funky groovy display font' },
    secondaryFont: { role: 'Supporting line', style: 'Clean uppercase sans serif', example: 'A simple geometric sans' },
    warnings: ['Two fonts max.', 'The groovy font is loud — pair it with a plain sans so it doesn\'t get noisy.'],
    hierarchyNote: 'The groovy word carries the arch; the sans line sits quietly beneath it.',
  },
  western: {
    name: 'Western Serif + Simple Uppercase',
    mainFont: { role: 'Hero word', style: 'Western slab / tuscan serif', example: 'A spurred western serif' },
    secondaryFont: { role: 'Supporting line', style: 'Simple condensed uppercase', example: 'A plain condensed sans' },
    warnings: ['Two fonts max.', 'Western serifs are decorative — keep the second font very plain for balance.'],
    hierarchyNote: 'Western serif leads; condensed uppercase supports along the badge edges.',
  },
  'soft-feminine': {
    name: 'Script + Clean Serif',
    mainFont: { role: 'Hero word', style: 'Elegant but readable script or rounded serif', example: 'A high-contrast calligraphy or soft serif' },
    secondaryFont: { role: 'Supporting line', style: 'Clean thin serif or sans', example: 'A light, airy serif' },
    warnings: ['Two fonts max.', 'If the script is very loopy, set the hero word in the serif instead so it stays readable.'],
    hierarchyNote: 'The script is the star only if it reads clearly; otherwise let the serif lead.',
  },
  kids: {
    name: 'Bubble + Rounded Sans',
    mainFont: { role: 'Hero word', style: 'Bubble / balloon display', example: 'A fat rounded bubble font' },
    secondaryFont: { role: 'Supporting line', style: 'Rounded sans serif', example: 'A friendly rounded sans' },
    warnings: ['Two fonts max.', 'Avoid thin fonts entirely — everything should feel chunky and friendly.'],
    hierarchyNote: 'Big bubble hero word; rounded sans for the smaller supporting words.',
  },
  minimal: {
    name: 'Bold Sans + Thin Sans',
    mainFont: { role: 'Hero word', style: 'Bold geometric sans serif', example: 'A heavy grotesque sans' },
    secondaryFont: { role: 'Supporting line', style: 'Thin / light sans, wide letter-spacing', example: 'The light weight of the same family' },
    warnings: ['Two fonts (or one family, two weights).', 'Contrast comes from weight and spacing, not decoration.'],
    hierarchyNote: 'Bold hero word, widely-spaced thin caption — restraint is the point.',
  },
  boho: {
    name: 'Rounded Serif + Airy Sans',
    mainFont: { role: 'Hero word', style: 'Soft rounded serif', example: 'A warm humanist serif' },
    secondaryFont: { role: 'Supporting line', style: 'Airy light sans, letter-spaced', example: 'A light minimal sans' },
    warnings: ['Two fonts max.', 'Keep both fonts calm; the earthy palette carries the character.'],
    hierarchyNote: 'Rounded serif hero word; the light sans whispers underneath.',
  },
  cozy: {
    name: 'Warm Serif + Handwritten',
    mainFont: { role: 'Hero word', style: 'Warm bold serif or slab', example: 'A cozy vintage serif' },
    secondaryFont: { role: 'Supporting line', style: 'Casual handwritten', example: 'A relaxed handwriting font' },
    warnings: ['Two fonts max.', 'Keep the handwritten line short so it stays legible at thumbnail size.'],
    hierarchyNote: 'Warm serif hero word up top; handwritten tagline for a homey touch.',
  },
  varsity: {
    name: 'Collegiate Block + Script',
    mainFont: { role: 'Hero words', style: 'Collegiate athletic block serif (arched)', example: 'A varsity / college block font' },
    secondaryFont: { role: 'Connector word', style: 'Retro script for the small “and/&”', example: 'A vintage script' },
    warnings: ['Two fonts max.', 'Keep the block letters even along the arch; the script is only for the tiny connector word.'],
    hierarchyNote: 'Big arched block words carry it; a small script word nestles between the lines.',
  },
  coquette: {
    name: 'Delicate Script + Line Caps',
    mainFont: { role: 'Hero phrase', style: 'Flowing delicate script', example: 'An airy signature script' },
    secondaryFont: { role: 'Supporting line', style: 'Light spaced serif caps', example: 'A thin all-caps serif' },
    warnings: ['Two fonts max.', 'Keep line weights thin and consistent so it stays dainty.'],
    hierarchyNote: 'The script leads; delicate line-art icons sit above it as the “title”.',
  },
  sketch: {
    name: 'Marker Brush + Doodle Caps',
    mainFont: { role: 'Hero words', style: 'Hand-drawn marker brush script', example: 'A loose brush marker font' },
    secondaryFont: { role: 'Alt words', style: 'Rough hand-drawn caps', example: 'A sketchy marker capital' },
    warnings: ['Two fonts max.', 'Everything should look hand-drawn — avoid any clean/geometric font here.'],
    hierarchyNote: 'Alternate brush script and rough caps line by line for the sketchbook feel.',
  },
  glam: {
    name: 'Bold Serif Caps + Brush Script',
    mainFont: { role: 'Hero words', style: 'Bold serif caps (one word leopard-filled)', example: 'A high-contrast serif' },
    secondaryFont: { role: 'Script word', style: 'Confident brush script', example: 'A glossy brush script' },
    warnings: ['Two fonts max.', 'Put print/pattern in only one word; keep the rest solid so it reads.'],
    hierarchyNote: 'Solid serif caps plus one patterned word; a brush script closes it out.',
  },
  edgy: {
    name: 'Grunge Brush + Checker Caps',
    mainFont: { role: 'Hero words', style: 'Grungy spray/brush display', example: 'A distressed brush font' },
    secondaryFont: { role: 'Alt words', style: 'Bold caps for checkerboard fill', example: 'A heavy blocky sans' },
    warnings: ['Two fonts max.', 'Mixing textures is the look — but keep the palette to pink, black and white.'],
    hierarchyNote: 'Loud brush words trade off with checkerboard-filled caps line by line.',
  },
  cartoon: {
    name: 'Rounded Marker + Simple Sans',
    mainFont: { role: 'Title (small)', style: 'Chunky rounded marker', example: 'A playful rounded marker font' },
    secondaryFont: { role: 'Labels', style: 'Simple friendly sans', example: 'A plain rounded sans' },
    warnings: ['Two fonts max.', 'Keep text minimal — the characters carry the design, not the type.'],
    hierarchyNote: 'A small friendly title sits above or below the row of cute characters.',
  },
  'bright-doodle': {
    name: 'Puffy Hand-Lettered + Sans',
    mainFont: { role: 'Hero words', style: 'Puffy hand-drawn bubble caps', example: 'A rounded marker/bubble font' },
    secondaryFont: { role: 'Tiny extras', style: 'Simple rounded sans', example: 'A friendly small sans' },
    warnings: ['Two fonts max.', 'Let colour do the work — one puffy lettering style, many colours.'],
    hierarchyNote: 'Big multicolour puffy words stacked; doodles fill the corners.',
  },
  watercolor: {
    name: 'Flowing Script + Serif Labels',
    mainFont: { role: 'Name / phrase', style: 'Thin flowing script', example: 'A delicate signature script' },
    secondaryFont: { role: 'Small labels', style: 'Simple serif caps', example: 'A light serif for book spines/labels' },
    warnings: ['Two fonts max.', 'Keep the script dark and thin so it reads over soft watercolour.'],
    hierarchyNote: 'Painterly art leads; a thin script name anchors the bottom.',
  },
}

// Optional accent font guidance when a design genuinely needs a third font.
export const ACCENT_FONT = {
  role: 'Accent word (only if needed)',
  style: 'A small decorative font for a single tiny word like "the" or "&"',
  example: 'A ligature script or a tiny serif',
}

// ---------------------------------------------------------------------------
// Object catalog (PRD Feature 3). Maps common object names to a default
// role and a natural rotation. The engine still re-ranks roles per design.
// ---------------------------------------------------------------------------

export interface ObjectMeta {
  role: ObjectRole
  rotation: string
  size: string
}

// Default fallbacks by role.
export const ROLE_DEFAULTS: Record<ObjectRole, ObjectMeta> = {
  main: { role: 'main', rotation: 'none (upright)', size: '~35–40% of the hero word width' },
  supporting: { role: 'supporting', rotation: '10–20 degrees', size: '~60–70% of the main object' },
  filler: { role: 'filler', rotation: 'varied, subtle', size: 'tiny — smaller than any supporting object' },
  frame: { role: 'frame', rotation: 'follows the frame curve', size: 'spans the edge it decorates' },
  background: { role: 'background', rotation: 'none', size: 'large but faded / low-contrast' },
}

// Known objects → default role + rotation. Names are matched loosely.
export const OBJECT_CATALOG: Record<string, ObjectMeta> = {
  // Classic "main object" candidates
  apple: { role: 'main', rotation: 'none', size: '~35% of the hero word width' },
  'coffee cup': { role: 'main', rotation: 'none', size: 'center-weight, ~40% of canvas height' },
  coffee: { role: 'main', rotation: 'none', size: 'center-weight' },
  mug: { role: 'main', rotation: 'none', size: 'center-weight' },
  cup: { role: 'main', rotation: 'none', size: 'center-weight' },
  heart: { role: 'filler', rotation: 'slight tilt', size: 'small' },
  book: { role: 'main', rotation: 'none', size: '~35% of hero word width' },
  books: { role: 'supporting', rotation: '5–10 degrees', size: '~60% of main object' },
  notebook: { role: 'supporting', rotation: '8–12 degrees', size: '~60% of main object' },
  pencil: { role: 'supporting', rotation: '15–20 degrees', size: 'thin, ~70% of main object length' },
  pen: { role: 'supporting', rotation: '15–20 degrees', size: 'thin' },
  crayon: { role: 'frame', rotation: '20–25 degrees', size: 'small' },
  ruler: { role: 'frame', rotation: '20–25 degrees', size: 'thin edge piece' },
  paperclip: { role: 'filler', rotation: 'varied', size: 'tiny' },
  star: { role: 'filler', rotation: 'varied', size: 'tiny' },
  stars: { role: 'filler', rotation: 'varied', size: 'tiny clusters' },
  sparkle: { role: 'filler', rotation: 'varied', size: 'tiny' },
  sparkles: { role: 'filler', rotation: 'varied', size: 'tiny' },
  dot: { role: 'filler', rotation: 'none', size: 'tiny' },
  dots: { role: 'filler', rotation: 'none', size: 'tiny' },
  flower: { role: 'supporting', rotation: '10–15 degrees', size: '~60% of main object' },
  flowers: { role: 'frame', rotation: 'follows the wreath curve', size: 'corner clusters' },
  floral: { role: 'frame', rotation: 'follows the wreath curve', size: 'corner clusters' },
  leaf: { role: 'filler', rotation: 'varied', size: 'small' },
  leaves: { role: 'frame', rotation: 'follows the wreath curve', size: 'trailing greenery' },
  bow: { role: 'supporting', rotation: 'none', size: '~50% of main object' },
  sun: { role: 'background', rotation: 'none', size: 'large arc behind the text' },
  rainbow: { role: 'background', rotation: 'none', size: 'arc behind the text' },
  cactus: { role: 'supporting', rotation: 'none', size: '~65% of main object' },
  cloud: { role: 'filler', rotation: 'none', size: 'small' },
  moon: { role: 'supporting', rotation: 'none', size: '~50% of main object' },
  butterfly: { role: 'filler', rotation: 'varied', size: 'small' },
  ghost: { role: 'main', rotation: 'none', size: 'center-weight' },
  pumpkin: { role: 'main', rotation: 'none', size: '~40% of hero word width' },
  snowflake: { role: 'filler', rotation: 'varied', size: 'tiny' },
  stethoscope: { role: 'main', rotation: 'slight curve', size: '~40% of hero word width' },
  syringe: { role: 'supporting', rotation: '15–20 degrees', size: 'thin' },
  cross: { role: 'filler', rotation: 'none', size: 'small' },
  // Western + coquette line-art motifs
  boot: { role: 'supporting', rotation: 'none', size: '~65% of main object' },
  boots: { role: 'supporting', rotation: 'none', size: '~70% of main object' },
  'cowboy boot': { role: 'supporting', rotation: 'none', size: '~65% of main object' },
  'cowboy hat': { role: 'main', rotation: 'slight tilt', size: '~40% of hero word width' },
  hat: { role: 'main', rotation: 'slight tilt', size: '~40% of hero word width' },
  horseshoe: { role: 'supporting', rotation: 'none', size: '~55% of main object' },
  bandana: { role: 'filler', rotation: 'none', size: 'small' },
  ribbon: { role: 'supporting', rotation: 'none', size: '~50% of main object' },
  // Storybook / whimsical + funny-pet motifs
  frog: { role: 'main', rotation: 'none', size: 'center-weight' },
  sword: { role: 'supporting', rotation: '10–15 degrees', size: 'thin, tall' },
  dog: { role: 'main', rotation: 'none', size: 'center-weight' },
  cat: { role: 'main', rotation: 'none', size: 'center-weight' },
  beer: { role: 'supporting', rotation: 'none', size: '~55% of main object' },
  sunglasses: { role: 'filler', rotation: 'slight tilt', size: 'small' },
  // Edgy / Y2K doodle motifs
  lightning: { role: 'filler', rotation: 'none', size: 'small' },
  bolt: { role: 'filler', rotation: 'none', size: 'small' },
  crown: { role: 'filler', rotation: 'slight tilt', size: 'small' },
  // Coffee/bookish extras
  'coffee mug': { role: 'main', rotation: 'none', size: 'center-weight' },
  glasses: { role: 'filler', rotation: 'none', size: 'small' },
  // Cartoon character + doodle motifs
  pickle: { role: 'main', rotation: 'slight tilt', size: 'center-weight character' },
  cucumber: { role: 'main', rotation: 'slight tilt', size: 'center-weight character' },
  jar: { role: 'supporting', rotation: 'none', size: '~70% of main object' },
  smiley: { role: 'filler', rotation: 'none', size: 'small' },
  'smiley face': { role: 'filler', rotation: 'none', size: 'small' },
  pennant: { role: 'supporting', rotation: 'slight tilt', size: '~45% of main object' },
  flag: { role: 'supporting', rotation: 'slight tilt', size: '~45% of main object' },
}

// ---------------------------------------------------------------------------
// Niche hints — suggested main object + default vibe when the user leaves
// style vague.
// ---------------------------------------------------------------------------

export interface NicheHint {
  keywords: string[]
  suggestedObjects: string[]
  defaultVibe: VibeId
}

export const NICHE_HINTS: NicheHint[] = [
  { keywords: ['teacher', 'teaching', 'school', 'classroom', 'educator'], suggestedObjects: ['apple', 'pencil', 'stars', 'notebook'], defaultVibe: 'cute-retro' },
  { keywords: ['nurse', 'nursing', 'medical', 'rn'], suggestedObjects: ['stethoscope', 'heart', 'stars'], defaultVibe: 'cute-retro' },
  { keywords: ['mom', 'mama', 'mother', 'mommy'], suggestedObjects: ['heart', 'flowers', 'stars'], defaultVibe: 'groovy-retro' },
  { keywords: ['coffee', 'caffeine', 'latte', 'espresso'], suggestedObjects: ['coffee cup', 'stars', 'heart'], defaultVibe: 'cozy' },
  { keywords: ['book', 'bookish', 'reader', 'reading', 'library'], suggestedObjects: ['book', 'flowers', 'stars'], defaultVibe: 'boho' },
  { keywords: ['faith', 'christian', 'jesus', 'blessed', 'church'], suggestedObjects: ['cross', 'flowers', 'sun'], defaultVibe: 'boho' },
  { keywords: ['sports', 'baseball', 'football', 'soccer', 'game day', 'mom life'], suggestedObjects: ['stars', 'heart'], defaultVibe: 'western' },
  { keywords: ['halloween', 'spooky', 'fall', 'autumn', 'pumpkin'], suggestedObjects: ['pumpkin', 'ghost', 'stars'], defaultVibe: 'cozy' },
  { keywords: ['christmas', 'holiday', 'winter', 'santa'], suggestedObjects: ['snowflake', 'leaves', 'stars'], defaultVibe: 'cozy' },
  { keywords: ['western', 'cowgirl', 'cowboy', 'rodeo', 'country'], suggestedObjects: ['cactus', 'stars'], defaultVibe: 'western' },
  { keywords: ['kids', 'birthday', 'toddler', 'baby'], suggestedObjects: ['stars', 'rainbow', 'cloud'], defaultVibe: 'kids' },
]

// Common little "stop words" so the hero-word picker skips them.
export const STOP_WORDS = new Set([
  'a', 'an', 'the', 'is', 'my', 'to', 'of', 'and', 'or', 'i', 'im', "i'm",
  'you', 'your', 'in', 'on', 'at', 'it', 'be', 'am', 'are', 'for', 'with',
])

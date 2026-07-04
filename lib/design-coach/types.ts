// Core types for Design Autopilot Coach — a design-recipe engine for
// beginner Etsy PNG sellers. Everything here is framework-agnostic and
// runs fully offline (no API required for the MVP).

export type LayoutId =
  | 'stacked-typography'
  | 'center-object'
  | 'badge-circle'
  | 'retro-arch'
  | 'three-frame'
  | 'wreath-corner'
  | 'minimal-typography'

export type ObjectRole =
  | 'main'
  | 'supporting'
  | 'filler'
  | 'frame'
  | 'background'

// What the user fills out (Screen 2 — Design Idea Form).
export interface DesignInput {
  phrase: string
  niche: string
  productType: string
  style: string
  objects: string[]
  colourVibe: string
  season: string
  targetBuyer: string
  notes: string
}

export interface LayoutRecipe {
  id: LayoutId
  name: string
  why: string
  mainTextArea: string
  smallTextArea: string
  mainObjectArea: string
  supportingObjectArea: string
  fillerArea: string
  keepEmpty: string
}

export interface TextHierarchy {
  // The full phrase split into a hero (largest) word plus the rest.
  heroWord: string
  leadText: string // words before the hero
  tailText: string // words after the hero
  optionalSubText: string
  notes: string
}

export interface PlacedObject {
  name: string
  role: ObjectRole
  placement: string
  size: string // relative sizing guidance
  rotation: string // degrees or "none"
}

export interface FontPairing {
  name: string
  mainFont: { role: string; style: string; example: string }
  secondaryFont: { role: string; style: string; example: string }
  accentFont?: { role: string; style: string; example: string }
  warnings: string[]
  hierarchyNote: string
}

export interface ColourSwatch {
  role: string
  name: string
  hex: string
  usage: string
}

export interface ColourRecipe {
  name: string
  swatches: ColourSwatch[]
  usageNotes: string[]
}

export interface ScoreBreakdown {
  overall: number // 0-100
  readability: number // 0-10
  colourHarmony: number // 0-10
  spacing: number // 0-10
  fontPairing: number // 0-10
  objectBalance: number // 0-10
  thumbnailReadability: number // 0-10
}

export interface DesignScore {
  before: ScoreBreakdown
  after: ScoreBreakdown
  mainIssues: string[]
}

// A compact "improved direction" (Feature 11 — Generate 3 Better Versions).
export interface DesignVariant {
  id: 'clean' | 'trendy' | 'playful'
  title: string
  tagline: string
  layout: string
  textPlacement: string
  objectPlacement: string
  fontPairing: string
  palette: string
  warning: string
  aiPrompt: string
}

// The full recipe (Screen 3 / PRD section 13 output template).
export interface DesignRecipe {
  input: DesignInput
  summary: string
  layout: LayoutRecipe
  canvas: { size: string; safeMargin: string; background: string }
  textHierarchy: TextHierarchy
  objects: PlacedObject[]
  fonts: FontPairing
  colours: ColourRecipe
  whiteSpaceNotes: string[]
  warnings: string[]
  score: DesignScore
  aiPrompt: string
  buildSteps: string[]
  variants: DesignVariant[]
}

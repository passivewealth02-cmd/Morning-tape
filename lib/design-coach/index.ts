export * from './types'
export { generateRecipe, resolveVibe, pickLayout } from './engine'
export { VIBES, NICHE_HINTS } from './knowledge'

import { NICHE_HINTS } from './knowledge'

/** Suggest starter objects for a niche (used to prefill the form). */
export function suggestObjectsForNiche(niche: string): string[] {
  const n = niche.toLowerCase().trim()
  const hint = NICHE_HINTS.find(h => h.keywords.some(k => n.includes(k)))
  return hint?.suggestedObjects ?? []
}

/** Example presets for the "try an example" button. */
export const EXAMPLE_INPUTS = [
  {
    label: 'Teacher shirt',
    input: {
      phrase: 'Teaching is my jam',
      niche: 'Teacher',
      productType: 'T-shirt PNG',
      style: 'Cute retro',
      objects: ['apple', 'pencil', 'stars', 'notebook'],
      colourVibe: 'warm retro',
      season: 'Back to school',
      targetBuyer: 'Elementary teachers',
      notes: '',
    },
  },
  {
    label: 'Coffee tumbler',
    input: {
      phrase: 'But first coffee',
      niche: 'Coffee',
      productType: 'Tumbler PNG',
      style: 'Cozy vintage',
      objects: ['coffee cup', 'stars', 'heart'],
      colourVibe: 'warm cozy',
      season: '',
      targetBuyer: 'Coffee lovers',
      notes: '',
    },
  },
  {
    label: 'Nurse badge',
    input: {
      phrase: 'Nurse life',
      niche: 'Nurse',
      productType: 'T-shirt PNG',
      style: 'Western badge',
      objects: ['stethoscope', 'heart', 'stars'],
      colourVibe: 'desert',
      season: '',
      targetBuyer: 'ER nurses',
      notes: '',
    },
  },
] as const

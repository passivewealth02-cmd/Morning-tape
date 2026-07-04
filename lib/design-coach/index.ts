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
  {
    label: 'Varsity books & coffee',
    input: {
      phrase: 'Books and Coffee',
      niche: 'Bookish',
      productType: 'T-shirt PNG',
      style: 'Varsity collegiate',
      objects: [],
      colourVibe: 'pink and mint pastel',
      season: '',
      targetBuyer: 'Book lovers',
      notes: 'Arched block letters with a distressed texture.',
    },
  },
  {
    label: 'Coquette Texas',
    input: {
      phrase: "She's from Texas",
      niche: 'Western',
      productType: 'Sticker',
      style: 'Coquette line art',
      objects: ['boot', 'cowboy hat', 'bow', 'horseshoe'],
      colourVibe: 'single-colour denim blue',
      season: '',
      targetBuyer: 'Cowgirls',
      notes: 'Delicate one-colour line icons in a row above a script.',
    },
  },
  {
    label: 'Sketch books & coffee',
    input: {
      phrase: 'Read Books and Drink Coffee',
      niche: 'Coffee',
      productType: 'T-shirt PNG',
      style: 'Sketch marker doodle',
      objects: ['book', 'coffee cup', 'heart'],
      colourVibe: 'orange and blue two-tone',
      season: '',
      targetBuyer: 'Readers',
      notes: 'Hand-drawn marker look, two colours only.',
    },
  },
  {
    label: 'Cartoon pickles',
    input: {
      phrase: 'Pickle Lover',
      niche: 'Funny food',
      productType: 'Sticker',
      style: 'Cartoon characters',
      objects: ['pickle', 'jar', 'cucumber', 'stars'],
      colourVibe: 'green cartoon',
      season: '',
      targetBuyer: 'Pickle fans',
      notes: 'Cute food characters with faces, arms and legs in a row.',
    },
  },
  {
    label: 'Bright best teacher',
    input: {
      phrase: 'Best Teacher Ever',
      niche: 'Teacher',
      productType: 'T-shirt PNG',
      style: 'Bright colorful puffy letters',
      objects: ['heart', 'flower', 'smiley', 'stars'],
      colourVibe: 'colorful with gold glitter',
      season: '',
      targetBuyer: 'Teachers',
      notes: 'Every letter a different bright colour with a gold glitter outline.',
    },
  },
  {
    label: 'Watercolor teacher name',
    input: {
      phrase: 'Mrs. Collins',
      niche: 'Teacher',
      productType: 'T-shirt PNG',
      style: 'Watercolor floral',
      objects: ['book', 'flowers', 'leaves', 'butterfly'],
      colourVibe: 'soft watercolor',
      season: '',
      targetBuyer: 'Teachers',
      notes: 'Personalized: flowers growing from stacked books, name in script.',
    },
  },
] as const

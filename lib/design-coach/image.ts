import 'server-only'

// Turns a design-recipe prompt into an actual PNG using an image model.
// Defaults to OpenAI's gpt-image-1, which supports transparent backgrounds
// (ideal for Etsy PNG designs). Configurable via env so any OpenAI-style
// image endpoint can be swapped in without code changes.
//
//   OPENAI_API_KEY        required to enable the feature
//   COACH_IMAGE_MODEL     optional, defaults to "gpt-image-1"
//   COACH_IMAGE_BASE_URL  optional, defaults to OpenAI's API
//
// When no key is configured the caller gets a clear, actionable message
// instead of a crash.

export type ImageResult =
  | { ok: true; dataUrl: string; model: string }
  | { ok: false; error: string; code: 'no_key' | 'provider_error' | 'bad_request' }

const DEFAULT_MODEL = 'gpt-image-1'
const DEFAULT_BASE_URL = 'https://api.openai.com/v1'

const MAX_REFERENCES = 10

export function imageGenerationConfigured(): boolean {
  return Boolean(process.env.OPENAI_API_KEY)
}

function dataUrlToBlob(dataUrl: string): Blob | null {
  const m = /^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/.exec(dataUrl)
  if (!m) return null
  try {
    return new Blob([Buffer.from(m[2], 'base64')], { type: m[1] })
  } catch {
    return null
  }
}

// `references` are data URLs of images the user uploaded as a visual brief.
// When present, we call the image-edit endpoint so the model uses them as
// style guidance; otherwise we do a plain text-to-image generation.
export async function generateDesignImage(prompt: string, references: string[] = []): Promise<ImageResult> {
  const key = process.env.OPENAI_API_KEY
  if (!key) {
    return {
      ok: false,
      code: 'no_key',
      error:
        'Image generation is not set up yet. Add an OPENAI_API_KEY environment variable (from platform.openai.com) to enable the “Generate image” button.',
    }
  }

  const cleaned = prompt?.trim()
  if (!cleaned || cleaned.length < 10) {
    return { ok: false, code: 'bad_request', error: 'Generate a recipe first — the prompt was empty.' }
  }

  const model = process.env.COACH_IMAGE_MODEL || DEFAULT_MODEL
  const baseUrl = (process.env.COACH_IMAGE_BASE_URL || DEFAULT_BASE_URL).replace(/\/$/, '')

  const refBlobs = references
    .slice(0, MAX_REFERENCES)
    .map(dataUrlToBlob)
    .filter((b): b is Blob => b !== null)

  try {
    let res: Response
    if (refBlobs.length > 0) {
      // Image-edit endpoint: uploaded references steer the generated design.
      const form = new FormData()
      form.append('model', model)
      form.append(
        'prompt',
        (cleaned + ' Use the uploaded reference images as visual guidance for the style, palette and layout.').slice(0, 4000),
      )
      form.append('size', '1024x1536')
      form.append('background', 'transparent')
      refBlobs.forEach((blob, i) => form.append('image[]', blob, `reference-${i + 1}.png`))
      res = await fetch(`${baseUrl}/images/edits`, {
        method: 'POST',
        headers: { Authorization: `Bearer ${key}` },
        body: form,
      })
    } else {
      res = await fetch(`${baseUrl}/images/generations`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Authorization: `Bearer ${key}`,
        },
        body: JSON.stringify({
          model,
          prompt: cleaned.slice(0, 4000),
          // Portrait, close to a shirt/print ratio.
          size: '1024x1536',
          // gpt-image-1 only: transparent PNG, exactly what print sellers want.
          background: 'transparent',
          n: 1,
        }),
      })
    }

    if (!res.ok) {
      let detail = `Image provider returned ${res.status}.`
      try {
        const body = (await res.json()) as { error?: { message?: string } }
        if (body?.error?.message) detail = body.error.message
      } catch {
        /* keep the status-based message */
      }
      return { ok: false, code: 'provider_error', error: detail }
    }

    const data = (await res.json()) as { data?: Array<{ b64_json?: string; url?: string }> }
    const first = data.data?.[0]
    if (first?.b64_json) {
      return { ok: true, model, dataUrl: `data:image/png;base64,${first.b64_json}` }
    }
    if (first?.url) {
      // Some models (e.g. dall-e-3) return a URL instead of base64.
      return { ok: true, model, dataUrl: first.url }
    }
    return { ok: false, code: 'provider_error', error: 'The image provider returned no image.' }
  } catch (err) {
    console.error('generateDesignImage failed:', err)
    return { ok: false, code: 'provider_error', error: 'Could not reach the image provider. Try again in a moment.' }
  }
}

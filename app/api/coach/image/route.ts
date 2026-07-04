import { NextRequest, NextResponse } from 'next/server'
import { generateDesignImage } from '@/lib/design-coach/image'

export const runtime = 'nodejs'
export const maxDuration = 60

export async function POST(request: NextRequest) {
  let prompt = ''
  let references: string[] = []
  try {
    const body = await request.json()
    prompt = typeof body?.prompt === 'string' ? body.prompt : ''
    if (Array.isArray(body?.references)) {
      references = body.references.filter((r: unknown): r is string => typeof r === 'string' && r.startsWith('data:image/')).slice(0, 10)
    }
  } catch {
    return NextResponse.json({ error: 'Invalid request body.' }, { status: 400 })
  }

  if (!prompt.trim()) {
    return NextResponse.json({ error: 'Generate a recipe first — the prompt was empty.' }, { status: 400 })
  }

  const result = await generateDesignImage(prompt, references)

  if (!result.ok) {
    // no_key / bad_request are the user's to fix (422); provider errors are 502.
    const status = result.code === 'provider_error' ? 502 : 422
    return NextResponse.json({ error: result.error, code: result.code }, { status })
  }

  return NextResponse.json({ dataUrl: result.dataUrl, model: result.model })
}

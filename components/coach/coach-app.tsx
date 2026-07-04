'use client'

import { useState, useMemo, useEffect } from 'react'
import {
  Sparkles, Wand2, Copy, Check, RotateCcw, Save, Trash2, Palette,
  Type, LayoutGrid, AlertTriangle, Wind, ListChecks, Star, Lightbulb,
  ImageIcon, Download, Loader2, RefreshCw, Upload, X, ImagePlus,
} from 'lucide-react'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Textarea } from '@/components/ui/textarea'
import { Label } from '@/components/ui/label'
import { Badge } from '@/components/ui/badge'
import { cn } from '@/lib/utils'
import { generateRecipe } from '@/lib/design-coach/engine'
import { EXAMPLE_INPUTS, suggestObjectsForNiche } from '@/lib/design-coach'
import type { DesignInput, DesignRecipe, ScoreBreakdown } from '@/lib/design-coach/types'

const EMPTY: DesignInput = {
  phrase: '', niche: '', productType: '', style: '', objects: [],
  colourVibe: '', season: '', targetBuyer: '', notes: '',
}

const ROLE_COLOR: Record<string, string> = {
  main: 'bg-orange-500/15 text-orange-700 dark:text-orange-300 border-orange-500/30',
  supporting: 'bg-teal-500/15 text-teal-700 dark:text-teal-300 border-teal-500/30',
  filler: 'bg-violet-500/15 text-violet-700 dark:text-violet-300 border-violet-500/30',
  frame: 'bg-emerald-500/15 text-emerald-700 dark:text-emerald-300 border-emerald-500/30',
  background: 'bg-slate-500/15 text-slate-700 dark:text-slate-300 border-slate-500/30',
}

interface SavedRecipe { id: string; label: string; input: DesignInput }

const MAX_REFS = 10

export function CoachApp() {
  const [form, setForm] = useState<DesignInput>(EMPTY)
  const [objectsText, setObjectsText] = useState('')
  const [recipe, setRecipe] = useState<DesignRecipe | null>(null)
  const [saved, setSaved] = useState<SavedRecipe[]>([])
  const [variantIdx, setVariantIdx] = useState(1) // default to "Trendy Etsy"
  const [refs, setRefs] = useState<string[]>([]) // reference images as data URLs

  useEffect(() => {
    try {
      const raw = localStorage.getItem('dac-saved')
      if (raw) setSaved(JSON.parse(raw))
    } catch { /* ignore */ }
  }, [])

  function persist(next: SavedRecipe[]) {
    setSaved(next)
    try { localStorage.setItem('dac-saved', JSON.stringify(next)) } catch { /* ignore */ }
  }

  function set<K extends keyof DesignInput>(key: K, value: DesignInput[K]) {
    setForm(f => ({ ...f, [key]: value }))
  }

  function parseObjects(text: string): string[] {
    return text.split(',').map(s => s.trim()).filter(Boolean)
  }

  function handleGenerate() {
    const input: DesignInput = { ...form, objects: parseObjects(objectsText) }
    setRecipe(generateRecipe(input))
    setVariantIdx(1)
    setTimeout(() => document.getElementById('recipe-output')?.scrollIntoView({ behavior: 'smooth', block: 'start' }), 50)
  }

  function loadExample(i: number) {
    const ex = EXAMPLE_INPUTS[i].input
    setForm({ ...ex, objects: [...ex.objects] })
    setObjectsText(ex.objects.join(', '))
    setRecipe(generateRecipe({ ...ex, objects: [...ex.objects] }))
    setVariantIdx(1)
    setTimeout(() => document.getElementById('recipe-output')?.scrollIntoView({ behavior: 'smooth', block: 'start' }), 50)
  }

  function reset() {
    setForm(EMPTY); setObjectsText(''); setRecipe(null)
  }

  function saveCurrent() {
    if (!recipe) return
    const label = recipe.input.phrase || 'Untitled design'
    const entry: SavedRecipe = { id: crypto.randomUUID(), label, input: recipe.input }
    persist([entry, ...saved].slice(0, 20))
  }

  function loadSaved(s: SavedRecipe) {
    setForm(s.input)
    setObjectsText(s.input.objects.join(', '))
    setRecipe(generateRecipe(s.input))
    setVariantIdx(1)
    setTimeout(() => document.getElementById('recipe-output')?.scrollIntoView({ behavior: 'smooth', block: 'start' }), 50)
  }

  const suggestions = useMemo(() => suggestObjectsForNiche(form.niche), [form.niche])
  const canGenerate = form.phrase.trim().length > 0

  return (
    <main className="min-h-screen bg-gradient-to-b from-orange-50/60 via-background to-background dark:from-orange-950/20">
      {/* Header */}
      <header className="border-b bg-background/70 backdrop-blur sticky top-0 z-10">
        <div className="mx-auto max-w-6xl px-4 py-4 flex items-center gap-3">
          <div className="flex h-9 w-9 items-center justify-center rounded-lg bg-gradient-to-br from-orange-500 to-pink-500 text-white shadow-sm">
            <Wand2 className="h-5 w-5" />
          </div>
          <div>
            <h1 className="text-lg font-bold leading-tight">Design Autopilot Coach</h1>
            <p className="text-xs text-muted-foreground leading-tight">Turn a rough idea into a clean, sellable PNG design recipe</p>
          </div>
        </div>
      </header>

      <div className="mx-auto max-w-6xl px-4 py-8 grid gap-8 lg:grid-cols-[minmax(0,380px)_minmax(0,1fr)]">
        {/* Form */}
        <div className="lg:sticky lg:top-24 lg:self-start space-y-4">
          <Section title="Design idea" icon={<Lightbulb className="h-4 w-4" />}>
            <p className="text-sm text-muted-foreground mb-4">
              Answer a few simple questions. The coach handles layout, colour, fonts and placement for you.
            </p>

            <Field label="Phrase" hint="The words on your design" required>
              <Input value={form.phrase} onChange={e => set('phrase', e.target.value)} placeholder="Teaching is my jam" />
            </Field>

            <div className="grid grid-cols-2 gap-3">
              <Field label="Niche">
                <Input value={form.niche} onChange={e => set('niche', e.target.value)} placeholder="Teacher" />
              </Field>
              <Field label="Product">
                <Input value={form.productType} onChange={e => set('productType', e.target.value)} placeholder="T-shirt PNG" />
              </Field>
            </div>

            <div className="grid grid-cols-2 gap-3">
              <Field label="Style / vibe">
                <Input value={form.style} onChange={e => set('style', e.target.value)} placeholder="Cute retro" />
              </Field>
              <Field label="Colour vibe">
                <Input value={form.colourVibe} onChange={e => set('colourVibe', e.target.value)} placeholder="Warm retro" />
              </Field>
            </div>

            <Field label="Objects" hint="Comma separated">
              <Input value={objectsText} onChange={e => setObjectsText(e.target.value)} placeholder="apple, pencil, stars, notebook" />
              {suggestions.length > 0 && (
                <div className="mt-2 flex flex-wrap items-center gap-1.5">
                  <span className="text-xs text-muted-foreground">Suggested:</span>
                  {suggestions.map(s => (
                    <button
                      key={s}
                      type="button"
                      onClick={() => setObjectsText(t => {
                        const cur = parseObjects(t)
                        return cur.includes(s) ? t : [...cur, s].join(', ')
                      })}
                      className="rounded-full border px-2 py-0.5 text-xs hover:bg-accent transition-colors"
                    >
                      + {s}
                    </button>
                  ))}
                </div>
              )}
            </Field>

            <div className="grid grid-cols-2 gap-3">
              <Field label="Season / occasion">
                <Input value={form.season} onChange={e => set('season', e.target.value)} placeholder="Back to school" />
              </Field>
              <Field label="Target buyer">
                <Input value={form.targetBuyer} onChange={e => set('targetBuyer', e.target.value)} placeholder="Teachers" />
              </Field>
            </div>

            <Field label="Notes" hint="Optional">
              <Textarea value={form.notes} onChange={e => set('notes', e.target.value)} placeholder="Anything else the coach should know" rows={2} />
            </Field>

            <div className="flex flex-wrap gap-2 pt-1">
              <Button onClick={handleGenerate} disabled={!canGenerate} className="gap-1.5">
                <Sparkles className="h-4 w-4" /> Generate recipe
              </Button>
              <Button variant="outline" onClick={reset} className="gap-1.5">
                <RotateCcw className="h-4 w-4" /> Reset
              </Button>
            </div>
          </Section>

          <ReferenceUploader refs={refs} setRefs={setRefs} />

          <Section title="Try an example" icon={<Star className="h-4 w-4" />}>
            <div className="flex flex-wrap gap-2">
              {EXAMPLE_INPUTS.map((ex, i) => (
                <Button key={ex.label} variant="secondary" size="sm" onClick={() => loadExample(i)}>{ex.label}</Button>
              ))}
            </div>
          </Section>

          {saved.length > 0 && (
            <Section title="Saved recipes" icon={<Save className="h-4 w-4" />}>
              <ul className="space-y-1.5">
                {saved.map(s => (
                  <li key={s.id} className="flex items-center gap-2">
                    <button onClick={() => loadSaved(s)} className="flex-1 text-left text-sm truncate hover:text-primary transition-colors">
                      {s.label}
                    </button>
                    <button onClick={() => persist(saved.filter(x => x.id !== s.id))} className="text-muted-foreground hover:text-destructive">
                      <Trash2 className="h-3.5 w-3.5" />
                    </button>
                  </li>
                ))}
              </ul>
            </Section>
          )}
        </div>

        {/* Output */}
        <div id="recipe-output">
          {recipe ? <RecipeOutput recipe={recipe} onSave={saveCurrent} variantIdx={variantIdx} setVariantIdx={setVariantIdx} refs={refs} /> : <EmptyState />}
        </div>
      </div>
    </main>
  )
}

// ---------------------------------------------------------------------------

// Reads an image file and returns a downscaled JPEG data URL, so 10 references
// stay a reasonable payload size to upload to the image model.
function fileToScaledDataUrl(file: File, maxDim = 1024): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onerror = () => reject(new Error('read failed'))
    reader.onload = () => {
      const img = new Image()
      img.onerror = () => reject(new Error('decode failed'))
      img.onload = () => {
        const scale = Math.min(1, maxDim / Math.max(img.width, img.height))
        const w = Math.round(img.width * scale)
        const h = Math.round(img.height * scale)
        const canvas = document.createElement('canvas')
        canvas.width = w; canvas.height = h
        const ctx = canvas.getContext('2d')
        if (!ctx) return reject(new Error('no canvas'))
        ctx.drawImage(img, 0, 0, w, h)
        resolve(canvas.toDataURL('image/jpeg', 0.85))
      }
      img.src = reader.result as string
    }
    reader.readAsDataURL(file)
  })
}

function ReferenceUploader({ refs, setRefs }: { refs: string[]; setRefs: React.Dispatch<React.SetStateAction<string[]>> }) {
  const [busy, setBusy] = useState(false)

  async function onFiles(files: FileList | null) {
    if (!files || files.length === 0) return
    setBusy(true)
    const room = MAX_REFS - refs.length
    const picked = Array.from(files).filter(f => f.type.startsWith('image/')).slice(0, room)
    const scaled: string[] = []
    for (const f of picked) {
      try { scaled.push(await fileToScaledDataUrl(f)) } catch { /* skip unreadable */ }
    }
    setRefs(prev => [...prev, ...scaled].slice(0, MAX_REFS))
    setBusy(false)
  }

  const full = refs.length >= MAX_REFS

  return (
    <Section title="Reference images" icon={<ImagePlus className="h-4 w-4" />}>
      <p className="text-sm text-muted-foreground mb-3">
        Add up to {MAX_REFS} images (styles you like, colours, layouts). They&apos;ll steer the generated image — a visual brief for how you want it designed.
      </p>

      <label
        className={cn(
          'flex flex-col items-center justify-center gap-1.5 rounded-lg border border-dashed p-5 text-center transition-colors',
          full ? 'cursor-not-allowed opacity-50' : 'cursor-pointer hover:border-primary hover:bg-accent/40',
        )}
      >
        <Upload className="h-5 w-5 text-muted-foreground" />
        <span className="text-sm font-medium">{full ? `Maximum ${MAX_REFS} reached` : busy ? 'Adding…' : 'Click to upload'}</span>
        <span className="text-xs text-muted-foreground">{refs.length}/{MAX_REFS} added · PNG or JPG</span>
        <input type="file" accept="image/*" multiple hidden disabled={full || busy} onChange={e => { onFiles(e.target.files); e.target.value = '' }} />
      </label>

      {refs.length > 0 && (
        <>
          <div className="mt-3 grid grid-cols-4 gap-2">
            {refs.map((src, i) => (
              <div key={i} className="group relative aspect-square overflow-hidden rounded-md border">
                {/* eslint-disable-next-line @next/next/no-img-element */}
                <img src={src} alt={`Reference ${i + 1}`} className="h-full w-full object-cover" />
                <button
                  type="button"
                  onClick={() => setRefs(prev => prev.filter((_, j) => j !== i))}
                  className="absolute right-0.5 top-0.5 flex h-5 w-5 items-center justify-center rounded-full bg-black/60 text-white opacity-0 transition-opacity group-hover:opacity-100"
                  aria-label={`Remove reference ${i + 1}`}
                >
                  <X className="h-3 w-3" />
                </button>
              </div>
            ))}
          </div>
          <button type="button" onClick={() => setRefs([])} className="mt-2 text-xs text-muted-foreground hover:text-destructive">
            Clear all
          </button>
        </>
      )}
    </Section>
  )
}

type ImageState =
  | { status: 'idle' }
  | { status: 'loading' }
  | { status: 'done'; url: string }
  | { status: 'error'; message: string; needsKey: boolean }

function ImagePanel({ prompt, phrase, refs }: { prompt: string; phrase: string; refs: string[] }) {
  const [state, setState] = useState<ImageState>({ status: 'idle' })

  async function run() {
    setState({ status: 'loading' })
    try {
      const res = await fetch('/api/coach/image', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ prompt, references: refs }),
      })
      const data = await res.json()
      if (!res.ok) {
        setState({ status: 'error', message: data?.error ?? 'Something went wrong.', needsKey: data?.code === 'no_key' })
        return
      }
      setState({ status: 'done', url: data.dataUrl })
    } catch {
      setState({ status: 'error', message: 'Could not reach the server. Try again.', needsKey: false })
    }
  }

  const fileName = (phrase || 'design').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 'design'

  return (
    <div className="rounded-xl border bg-card p-5 shadow-sm">
      <div className="mb-3 flex flex-wrap items-start justify-between gap-2">
        <div>
          <div className="flex items-center gap-2 font-semibold"><ImageIcon className="h-4 w-4 text-pink-500" />Generate the image</div>
          <p className="mt-1 text-sm text-muted-foreground">
            Render this recipe as a transparent-background PNG you can download.
            {refs.length > 0 && <> Using <span className="font-medium text-foreground">{refs.length} reference image{refs.length > 1 ? 's' : ''}</span> to steer the style.</>}
          </p>
        </div>
        {state.status !== 'loading' && (
          <Button size="sm" className="gap-1.5" onClick={run}>
            {state.status === 'done' ? <RefreshCw className="h-4 w-4" /> : <Sparkles className="h-4 w-4" />}
            {state.status === 'done' ? 'Regenerate' : 'Generate image'}
          </Button>
        )}
      </div>

      {state.status === 'idle' && (
        <div className="flex min-h-[140px] flex-col items-center justify-center rounded-lg border border-dashed bg-muted/20 p-6 text-center">
          <ImageIcon className="mb-2 h-6 w-6 text-muted-foreground" />
          <p className="text-sm text-muted-foreground">Click <span className="font-medium text-foreground">Generate image</span> to turn the recipe into artwork.</p>
        </div>
      )}

      {state.status === 'loading' && (
        <div className="flex min-h-[220px] flex-col items-center justify-center rounded-lg border bg-muted/20 p-6 text-center">
          <Loader2 className="mb-3 h-7 w-7 animate-spin text-primary" />
          <p className="text-sm font-medium">Designing your PNG…</p>
          <p className="mt-1 text-xs text-muted-foreground">This usually takes 10–25 seconds.</p>
        </div>
      )}

      {state.status === 'done' && (
        <div className="space-y-3">
          <div className="overflow-hidden rounded-lg border bg-[repeating-conic-gradient(#0000000a_0_25%,transparent_0_50%)] bg-[length:20px_20px] p-3">
            {/* eslint-disable-next-line @next/next/no-img-element */}
            <img src={state.url} alt={`Generated design for “${phrase}”`} className="mx-auto max-h-[560px] w-auto rounded" />
          </div>
          <div className="flex flex-wrap gap-2">
            <a href={state.url} download={`${fileName}.png`}>
              <Button size="sm" variant="outline" className="gap-1.5"><Download className="h-4 w-4" /> Download PNG</Button>
            </a>
          </div>
        </div>
      )}

      {state.status === 'error' && (
        <div className="rounded-lg border border-amber-500/30 bg-amber-500/5 p-4">
          <div className="flex gap-2 text-sm">
            <AlertTriangle className="mt-0.5 h-4 w-4 shrink-0 text-amber-500" />
            <div>
              <p className="text-foreground/90">{state.message}</p>
              {state.needsKey && (
                <p className="mt-1 text-xs text-muted-foreground">
                  You can still copy the AI prompt below and paste it into any image tool (ChatGPT, Ideogram, Midjourney).
                </p>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

function EmptyState() {
  return (
    <div className="flex min-h-[420px] flex-col items-center justify-center rounded-xl border border-dashed bg-card/40 p-10 text-center">
      <div className="mb-4 flex h-14 w-14 items-center justify-center rounded-2xl bg-gradient-to-br from-orange-500 to-pink-500 text-white">
        <Wand2 className="h-7 w-7" />
      </div>
      <h2 className="text-xl font-bold">Your design recipe appears here</h2>
      <p className="mt-2 max-w-md text-sm text-muted-foreground">
        Fill out the idea form and hit <span className="font-medium text-foreground">Generate recipe</span>. You&apos;ll get a layout,
        object placement map, font pairing, colour palette with roles, white-space notes, a design score, warnings, and a copy-paste AI prompt.
      </p>
      <p className="mt-4 text-xs text-muted-foreground">Or load one of the examples on the left to see it in action.</p>
    </div>
  )
}

function RecipeOutput({
  recipe, onSave, variantIdx, setVariantIdx, refs,
}: {
  recipe: DesignRecipe
  onSave: () => void
  variantIdx: number
  setVariantIdx: (i: number) => void
  refs: string[]
}) {
  const { input, layout, canvas, textHierarchy: hero, objects, fonts, colours, score } = recipe
  const [savedFlash, setSavedFlash] = useState(false)

  return (
    <div className="space-y-5">
      {/* Summary + score header */}
      <div className="rounded-xl border bg-card p-5 shadow-sm">
        <div className="flex flex-wrap items-start justify-between gap-3">
          <div>
            <Badge variant="secondary" className="mb-2">{layout.name}</Badge>
            <h2 className="text-xl font-bold leading-snug">{input.phrase || 'Your design'}</h2>
            <p className="mt-1 max-w-xl text-sm text-muted-foreground">{recipe.summary}</p>
          </div>
          <div className="flex flex-col items-center rounded-lg border bg-background px-4 py-2">
            <span className="text-xs text-muted-foreground">Potential</span>
            <span className="text-3xl font-black tabular-nums text-emerald-600 dark:text-emerald-400">{score.after.overall}</span>
            <span className="text-[10px] text-muted-foreground">from {score.before.overall} today</span>
          </div>
        </div>
        <div className="mt-4 flex flex-wrap gap-2">
          <Button size="sm" variant="outline" className="gap-1.5" onClick={() => { onSave(); setSavedFlash(true); setTimeout(() => setSavedFlash(false), 1500) }}>
            {savedFlash ? <Check className="h-4 w-4 text-emerald-500" /> : <Save className="h-4 w-4" />}
            {savedFlash ? 'Saved' : 'Save recipe'}
          </Button>
        </div>
      </div>

      {/* Generate image */}
      <ImagePanel prompt={recipe.aiPrompt} phrase={input.phrase} refs={refs} />

      {/* Score breakdown */}
      <Panel icon={<Star className="h-4 w-4" />} title="Design score" subtitle="Typical first attempt vs. this recipe applied">
        <div className="grid gap-x-6 gap-y-3 sm:grid-cols-2">
          <ScoreRow label="Readability" before={score.before.readability} after={score.after.readability} />
          <ScoreRow label="Colour harmony" before={score.before.colourHarmony} after={score.after.colourHarmony} />
          <ScoreRow label="Spacing" before={score.before.spacing} after={score.after.spacing} />
          <ScoreRow label="Font pairing" before={score.before.fontPairing} after={score.after.fontPairing} />
          <ScoreRow label="Object balance" before={score.before.objectBalance} after={score.after.objectBalance} />
          <ScoreRow label="Thumbnail readability" before={score.before.thumbnailReadability} after={score.after.thumbnailReadability} />
        </div>
        <div className="mt-4 rounded-lg bg-muted/50 p-3">
          <p className="mb-1.5 text-xs font-semibold uppercase tracking-wide text-muted-foreground">Main things to fix</p>
          <ul className="space-y-1 text-sm">
            {score.mainIssues.map((m, i) => <li key={i} className="flex gap-2"><span className="text-amber-500">•</span>{m}</li>)}
          </ul>
        </div>
      </Panel>

      {/* Layout + canvas */}
      <Panel icon={<LayoutGrid className="h-4 w-4" />} title="Recommended layout" subtitle={layout.why}>
        <div className="grid gap-3 sm:grid-cols-2">
          <PlacementItem label="Hero text" value={layout.mainTextArea} />
          <PlacementItem label="Smaller text" value={layout.smallTextArea} />
          <PlacementItem label="Main object" value={layout.mainObjectArea} />
          <PlacementItem label="Supporting objects" value={layout.supportingObjectArea} />
          <PlacementItem label="Filler" value={layout.fillerArea} />
          <PlacementItem label="Keep empty" value={layout.keepEmpty} />
        </div>
        <div className="mt-4 flex flex-wrap gap-2 text-xs">
          <Chip label="Canvas" value={canvas.size} />
          <Chip label="Margin" value={canvas.safeMargin} />
          <Chip label="Background" value={canvas.background} />
        </div>
      </Panel>

      {/* Text hierarchy */}
      <Panel icon={<Type className="h-4 w-4" />} title="Text placement" subtitle={hero.notes}>
        <div className="flex flex-col items-center justify-center gap-1 rounded-lg border bg-muted/30 py-6 text-center">
          {hero.leadText && <span className="text-sm text-muted-foreground">{hero.leadText}</span>}
          <span className="text-3xl font-black tracking-tight">{hero.heroWord}</span>
          {hero.tailText && <span className="text-sm text-muted-foreground">{hero.tailText}</span>}
          {hero.optionalSubText && <span className="mt-1 text-xs italic text-muted-foreground">{hero.optionalSubText}</span>}
        </div>
      </Panel>

      {/* Objects */}
      {objects.length > 0 && (
        <Panel icon={<LayoutGrid className="h-4 w-4" />} title="Object placement map" subtitle="Each object gets a role, a spot, a size and an angle">
          <ul className="space-y-2">
            {objects.map((o, i) => (
              <li key={i} className="flex flex-col gap-1 rounded-lg border p-3 sm:flex-row sm:items-center sm:gap-3">
                <div className="flex items-center gap-2 sm:w-44 sm:shrink-0">
                  <span className={cn('rounded-full border px-2 py-0.5 text-[10px] font-semibold uppercase tracking-wide', ROLE_COLOR[o.role])}>{o.role}</span>
                  <span className="font-medium capitalize">{o.name}</span>
                </div>
                <div className="text-sm text-muted-foreground">
                  {o.placement}. <span className="text-foreground/70">Size {o.size}.</span>{' '}
                  {o.rotation !== 'none' && o.rotation !== 'none (upright)' && <span className="text-foreground/70">Rotate {o.rotation}.</span>}
                </div>
              </li>
            ))}
          </ul>
        </Panel>
      )}

      {/* Fonts */}
      <Panel icon={<Type className="h-4 w-4" />} title="Font pairing" subtitle={`${fonts.name} — ${fonts.hierarchyNote}`}>
        <div className="grid gap-3 sm:grid-cols-2">
          <FontCard role={fonts.mainFont.role} style={fonts.mainFont.style} example={fonts.mainFont.example} primary />
          <FontCard role={fonts.secondaryFont.role} style={fonts.secondaryFont.style} example={fonts.secondaryFont.example} />
        </div>
        <ul className="mt-3 space-y-1 text-sm text-muted-foreground">
          {fonts.warnings.map((w, i) => <li key={i} className="flex gap-2"><span className="text-amber-500">•</span>{w}</li>)}
        </ul>
      </Panel>

      {/* Colours */}
      <Panel icon={<Palette className="h-4 w-4" />} title={`Colour recipe — ${colours.name}`} subtitle="Every colour has a role and a place to go">
        <div className="grid grid-cols-2 gap-2 sm:grid-cols-3">
          {colours.swatches.map((s, i) => (
            <div key={i} className="flex items-center gap-2.5 rounded-lg border p-2">
              <span className="h-9 w-9 shrink-0 rounded-md border" style={{ backgroundColor: s.hex }} />
              <div className="min-w-0">
                <p className="truncate text-xs font-semibold">{s.name}</p>
                <p className="truncate text-[10px] uppercase tracking-wide text-muted-foreground">{s.role}</p>
              </div>
            </div>
          ))}
        </div>
        <ul className="mt-3 space-y-1 text-sm text-muted-foreground">
          {colours.swatches.map((s, i) => <li key={i} className="flex gap-2"><span className="font-medium text-foreground">{s.name}:</span>{s.usage}</li>)}
        </ul>
        <ul className="mt-3 space-y-1 border-t pt-3 text-sm text-muted-foreground">
          {colours.usageNotes.map((n, i) => <li key={i} className="flex gap-2"><span className="text-amber-500">•</span>{n}</li>)}
        </ul>
      </Panel>

      {/* White space */}
      <Panel icon={<Wind className="h-4 w-4" />} title="White space notes">
        <ul className="space-y-1.5 text-sm">
          {recipe.whiteSpaceNotes.map((n, i) => <li key={i} className="flex gap-2 text-muted-foreground"><span className="text-sky-500">•</span>{n}</li>)}
        </ul>
      </Panel>

      {/* Warnings */}
      <Panel icon={<AlertTriangle className="h-4 w-4 text-amber-500" />} title="Design warnings" subtitle="Common beginner mistakes to avoid">
        <ul className="space-y-1.5 text-sm">
          {recipe.warnings.map((w, i) => (
            <li key={i} className="flex gap-2 rounded-md bg-amber-500/5 px-2 py-1.5 text-foreground/90">
              <AlertTriangle className="mt-0.5 h-3.5 w-3.5 shrink-0 text-amber-500" />{w}
            </li>
          ))}
        </ul>
      </Panel>

      {/* AI prompt */}
      <CopyBlock title="AI prompt" subtitle="Paste into your AI design tool" text={recipe.aiPrompt} />

      {/* Build steps */}
      <Panel icon={<ListChecks className="h-4 w-4" />} title="Canva / Kittl build steps">
        <ol className="space-y-2">
          {recipe.buildSteps.map((s, i) => (
            <li key={i} className="flex gap-3 text-sm">
              <span className="flex h-5 w-5 shrink-0 items-center justify-center rounded-full bg-primary/10 text-xs font-bold text-primary">{i + 1}</span>
              <span className="pt-0.5">{s}</span>
            </li>
          ))}
        </ol>
      </Panel>

      {/* Variants */}
      <Panel icon={<Sparkles className="h-4 w-4" />} title="3 better versions" subtitle="Same idea, three directions to pick from">
        <div className="mb-3 flex flex-wrap gap-2">
          {recipe.variants.map((v, i) => (
            <button
              key={v.id}
              onClick={() => setVariantIdx(i)}
              className={cn(
                'rounded-full border px-3 py-1 text-sm font-medium transition-colors',
                variantIdx === i ? 'border-primary bg-primary text-primary-foreground' : 'hover:bg-accent',
              )}
            >
              {['A', 'B', 'C'][i]} · {v.title}
            </button>
          ))}
        </div>
        {recipe.variants[variantIdx] && (
          <div className="space-y-3">
            <p className="text-sm italic text-muted-foreground">{recipe.variants[variantIdx].tagline}</p>
            <div className="grid gap-2 sm:grid-cols-2">
              <MiniRow label="Layout" value={recipe.variants[variantIdx].layout} />
              <MiniRow label="Fonts" value={recipe.variants[variantIdx].fontPairing} />
              <MiniRow label="Text" value={recipe.variants[variantIdx].textPlacement} />
              <MiniRow label="Objects" value={recipe.variants[variantIdx].objectPlacement} />
              <MiniRow label="Palette" value={recipe.variants[variantIdx].palette} />
              <MiniRow label="Watch out" value={recipe.variants[variantIdx].warning} />
            </div>
            <CopyBlock title="AI prompt for this version" text={recipe.variants[variantIdx].aiPrompt} compact />
          </div>
        )}
      </Panel>
    </div>
  )
}

// ---------------------------------------------------------------------------
// Small presentational pieces
// ---------------------------------------------------------------------------

function Section({ title, icon, children }: { title: string; icon: React.ReactNode; children: React.ReactNode }) {
  return (
    <div className="rounded-xl border bg-card p-5 shadow-sm">
      <div className="mb-3 flex items-center gap-2 text-sm font-semibold">
        <span className="text-primary">{icon}</span>{title}
      </div>
      {children}
    </div>
  )
}

function Panel({ icon, title, subtitle, children }: { icon: React.ReactNode; title: string; subtitle?: string; children: React.ReactNode }) {
  return (
    <div className="rounded-xl border bg-card p-5 shadow-sm">
      <div className="mb-3">
        <div className="flex items-center gap-2 font-semibold">{icon}{title}</div>
        {subtitle && <p className="mt-1 text-sm text-muted-foreground">{subtitle}</p>}
      </div>
      {children}
    </div>
  )
}

function Field({ label, hint, required, children }: { label: string; hint?: string; required?: boolean; children: React.ReactNode }) {
  return (
    <div className="mb-3 space-y-1.5">
      <Label className="flex items-center gap-1.5 text-xs">
        {label}
        {required && <span className="text-destructive">*</span>}
        {hint && <span className="font-normal text-muted-foreground">— {hint}</span>}
      </Label>
      {children}
    </div>
  )
}

function ScoreRow({ label, before, after }: { label: string; before: number; after: number }) {
  return (
    <div>
      <div className="mb-1 flex items-center justify-between text-sm">
        <span>{label}</span>
        <span className="tabular-nums text-muted-foreground">{before} → <span className="font-semibold text-emerald-600 dark:text-emerald-400">{after}</span>/10</span>
      </div>
      <div className="relative h-2 overflow-hidden rounded-full bg-muted">
        <div className="absolute inset-y-0 left-0 rounded-full bg-muted-foreground/30" style={{ width: `${before * 10}%` }} />
        <div className="absolute inset-y-0 left-0 rounded-full bg-emerald-500" style={{ width: `${after * 10}%`, opacity: 0.85 }} />
      </div>
    </div>
  )
}

function PlacementItem({ label, value }: { label: string; value: string }) {
  return (
    <div className="rounded-lg border bg-muted/20 p-3">
      <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">{label}</p>
      <p className="mt-0.5 text-sm">{value}</p>
    </div>
  )
}

function Chip({ label, value }: { label: string; value: string }) {
  return (
    <span className="inline-flex items-center gap-1.5 rounded-full border bg-background px-2.5 py-1">
      <span className="font-semibold text-muted-foreground">{label}</span>
      <span>{value}</span>
    </span>
  )
}

function FontCard({ role, style, example, primary }: { role: string; style: string; example: string; primary?: boolean }) {
  return (
    <div className={cn('rounded-lg border p-3', primary && 'border-primary/40 bg-primary/5')}>
      <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">{role}</p>
      <p className="mt-1 font-semibold">{style}</p>
      <p className="mt-0.5 text-xs text-muted-foreground">Try: {example}</p>
    </div>
  )
}

function MiniRow({ label, value }: { label: string; value: string }) {
  return (
    <div className="rounded-lg border bg-muted/20 p-2.5">
      <p className="text-[10px] font-semibold uppercase tracking-wide text-muted-foreground">{label}</p>
      <p className="mt-0.5 text-sm">{value}</p>
    </div>
  )
}

function CopyBlock({ title, subtitle, text, compact }: { title: string; subtitle?: string; text: string; compact?: boolean }) {
  const [copied, setCopied] = useState(false)
  function copy() {
    navigator.clipboard?.writeText(text).then(() => { setCopied(true); setTimeout(() => setCopied(false), 1500) })
  }
  return (
    <div className={cn('rounded-xl border bg-card shadow-sm', compact ? 'p-4' : 'p-5')}>
      <div className="mb-2 flex items-center justify-between gap-2">
        <div>
          <div className="flex items-center gap-2 font-semibold"><Sparkles className="h-4 w-4 text-pink-500" />{title}</div>
          {subtitle && <p className="text-sm text-muted-foreground">{subtitle}</p>}
        </div>
        <Button size="sm" variant="outline" className="gap-1.5" onClick={copy}>
          {copied ? <Check className="h-4 w-4 text-emerald-500" /> : <Copy className="h-4 w-4" />}
          {copied ? 'Copied' : 'Copy'}
        </Button>
      </div>
      <p className="whitespace-pre-wrap rounded-lg bg-muted/50 p-3 font-mono text-xs leading-relaxed">{text}</p>
    </div>
  )
}

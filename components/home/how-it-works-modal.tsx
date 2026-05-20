'use client'

import { useState, useEffect, useRef } from 'react'
import {
  X, ArrowRight, ArrowLeft, Zap, Users, ClipboardList, Bell, BarChart3,
  CheckCircle2, Play, Pause, Sparkles, Inbox, Send, ShieldCheck,
} from 'lucide-react'
import { MaintenaMark } from '@/components/brand/logo'

/* ---------- accent palettes (literal classes so Tailwind keeps them) ---------- */
type Accent = {
  iconBg: string; iconText: string; bar: string; pill: string; pillText: string; dot: string; ring: string; soft: string
}
const ACCENTS: Record<string, Accent> = {
  indigo:  { iconBg: 'bg-indigo-50',  iconText: 'text-indigo-600',  bar: 'from-indigo-500 to-violet-500',   pill: 'bg-indigo-100',  pillText: 'text-indigo-700',  dot: 'bg-indigo-600',  ring: 'ring-indigo-200',  soft: 'bg-indigo-50' },
  violet:  { iconBg: 'bg-violet-50',  iconText: 'text-violet-600',  bar: 'from-violet-500 to-fuchsia-500',  pill: 'bg-violet-100',  pillText: 'text-violet-700',  dot: 'bg-violet-600',  ring: 'ring-violet-200',  soft: 'bg-violet-50' },
  blue:    { iconBg: 'bg-blue-50',    iconText: 'text-blue-600',    bar: 'from-blue-500 to-cyan-500',       pill: 'bg-blue-100',    pillText: 'text-blue-700',    dot: 'bg-blue-600',    ring: 'ring-blue-200',    soft: 'bg-blue-50' },
  emerald: { iconBg: 'bg-emerald-50', iconText: 'text-emerald-600', bar: 'from-emerald-500 to-teal-500',    pill: 'bg-emerald-100', pillText: 'text-emerald-700', dot: 'bg-emerald-600', ring: 'ring-emerald-200', soft: 'bg-emerald-50' },
  amber:   { iconBg: 'bg-amber-50',   iconText: 'text-amber-600',   bar: 'from-amber-500 to-orange-500',    pill: 'bg-amber-100',   pillText: 'text-amber-700',   dot: 'bg-amber-500',   ring: 'ring-amber-200',   soft: 'bg-amber-50' },
  rose:    { iconBg: 'bg-rose-50',    iconText: 'text-rose-600',    bar: 'from-rose-500 to-pink-500',       pill: 'bg-rose-100',    pillText: 'text-rose-700',    dot: 'bg-rose-600',    ring: 'ring-rose-200',    soft: 'bg-rose-50' },
}

const enter = 'animate-in fade-in slide-in-from-bottom-2 fill-mode-both'

/* ============================ INTERACTIVE AI DEMO ============================ */
type Analysis = {
  category: string; trade: string; urgency: string; urgencyTone: string
  summary: string; eta: string
}
const SAMPLES = [
  'Water is leaking under my kitchen sink and pooling on the floor',
  'The heater stopped working and my apartment is freezing',
  'The outlet in the bedroom is sparking when I plug things in',
  'My front door lock is jammed and I cannot get in',
]
function classify(text: string): Analysis {
  const t = text.toLowerCase()
  if (/leak|water|sink|pipe|flood|drip|toilet|faucet|drain|burst/.test(t))
    return { category: 'Plumbing', trade: 'Licensed Plumber', urgency: 'High', urgencyTone: 'bg-red-100 text-red-700', summary: 'Active water leak with flooding risk — dispatch same day to prevent property damage.', eta: 'Same day' }
  if (/heat|hvac|furnace|\bac\b|air condition|cold|cooling|thermostat|freezing|no heat/.test(t))
    return { category: 'HVAC', trade: 'HVAC Technician', urgency: 'High', urgencyTone: 'bg-orange-100 text-orange-700', summary: 'Heating failure in an occupied unit — habitability concern, prioritize dispatch.', eta: 'Within 24h' }
  if (/power|outlet|electric|spark|breaker|wiring|shock|light|fuse/.test(t))
    return { category: 'Electrical', trade: 'Licensed Electrician', urgency: 'Urgent', urgencyTone: 'bg-red-100 text-red-700', summary: 'Sparking outlet poses a fire/safety hazard — flagged urgent, dispatch immediately.', eta: 'Immediate' }
  if (/lock|door|window|key|jam|hinge|broke|broken|glass/.test(t))
    return { category: 'General Repair', trade: 'Handyman / Locksmith', urgency: 'Medium', urgencyTone: 'bg-amber-100 text-amber-700', summary: 'Access/security issue — schedule promptly, tenant may be locked out.', eta: 'Within 48h' }
  if (/pest|roach|mice|rat|bug|ant|infest|cockroach/.test(t))
    return { category: 'Pest Control', trade: 'Pest Control Specialist', urgency: 'Medium', urgencyTone: 'bg-amber-100 text-amber-700', summary: 'Pest report — schedule inspection and treatment.', eta: 'Within 48h' }
  return { category: 'General', trade: 'Handyman', urgency: 'Low', urgencyTone: 'bg-gray-100 text-gray-600', summary: 'Standard maintenance request — route to the next available vendor.', eta: 'This week' }
}

function AiTriageDemo({ accent }: { accent: Accent }) {
  const [text, setText] = useState(SAMPLES[0])
  const [phase, setPhase] = useState<'idle' | 'analyzing' | 'done'>('idle')
  const [result, setResult] = useState<Analysis | null>(null)
  const timer = useRef<ReturnType<typeof setTimeout> | null>(null)

  const run = () => {
    if (!text.trim()) return
    setPhase('analyzing')
    setResult(null)
    if (timer.current) clearTimeout(timer.current)
    timer.current = setTimeout(() => {
      setResult(classify(text))
      setPhase('done')
    }, 1500)
  }

  useEffect(() => {
    run()
    return () => { if (timer.current) clearTimeout(timer.current) }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  return (
    <div className="rounded-xl border border-gray-200 bg-white p-4">
      <div className="flex items-center gap-1.5 mb-2">
        <Sparkles className={`w-3.5 h-3.5 ${accent.iconText}`} />
        <span className="text-[11px] font-semibold text-gray-700">Try it live — type a maintenance issue</span>
      </div>
      <textarea
        value={text}
        onChange={e => { setText(e.target.value); setPhase('idle') }}
        rows={2}
        className="w-full resize-none rounded-lg border border-gray-200 px-3 py-2 text-xs text-gray-700 focus:outline-none focus:ring-2 focus:ring-indigo-200"
        placeholder="e.g. There's a leak under my bathroom sink..."
      />
      <div className="flex flex-wrap gap-1.5 mt-2">
        {SAMPLES.map((s, i) => (
          <button
            key={i}
            onClick={() => { setText(s); setPhase('idle') }}
            className="text-[10px] px-2 py-1 rounded-full bg-gray-100 text-gray-600 hover:bg-gray-200 transition-colors"
          >
            {s.split(' ').slice(0, 3).join(' ')}…
          </button>
        ))}
      </div>
      <button
        onClick={run}
        disabled={phase === 'analyzing'}
        className="mt-3 w-full flex items-center justify-center gap-1.5 rounded-lg bg-gray-900 text-white text-xs font-medium py-2 hover:bg-gray-700 transition-colors disabled:opacity-60"
      >
        {phase === 'analyzing'
          ? (<><span className="w-3 h-3 border-2 border-white/40 border-t-white rounded-full animate-spin" /> Analyzing with AI…</>)
          : (<><Sparkles className="w-3.5 h-3.5" /> Analyze with AI</>)}
      </button>

      <div className="mt-3 min-h-[132px]">
        {phase === 'analyzing' && (
          <div className="grid grid-cols-2 gap-2">
            {Array.from({ length: 4 }).map((_, i) => (
              <div key={i} className="h-12 rounded-lg bg-gray-100 animate-pulse" style={{ animationDelay: `${i * 80}ms` }} />
            ))}
          </div>
        )}
        {phase === 'done' && result && (
          <div className="space-y-2">
            <div className="grid grid-cols-2 gap-2">
              {[
                { label: 'Category', value: result.category },
                { label: 'Recommended trade', value: result.trade },
                { label: 'Urgency', value: result.urgency, tone: result.urgencyTone },
                { label: 'Target ETA', value: result.eta },
              ].map((f, i) => (
                <div key={f.label} className={`rounded-lg ${accent.soft} px-2.5 py-1.5 ${enter}`} style={{ animationDelay: `${i * 70}ms` }}>
                  <div className="text-[9px] font-medium uppercase tracking-wide text-gray-400">{f.label}</div>
                  {f.tone
                    ? <span className={`inline-block mt-0.5 text-[10px] font-semibold px-1.5 py-0.5 rounded-full ${f.tone}`}>{f.value}</span>
                    : <div className="text-[11px] font-semibold text-gray-800">{f.value}</div>}
                </div>
              ))}
            </div>
            <div className={`rounded-lg border border-gray-100 bg-gray-50 px-3 py-2 ${enter}`} style={{ animationDelay: '300ms' }}>
              <div className="text-[9px] font-medium uppercase tracking-wide text-gray-400 mb-0.5">AI summary</div>
              <p className="text-[11px] text-gray-600 leading-relaxed">{result.summary}</p>
            </div>
          </div>
        )}
      </div>
    </div>
  )
}

/* ============================ VISUALS ============================ */
function CaptureVisual() {
  const emails = [
    { from: 'sarah@apt4b.com', subject: 'Leak under kitchen sink' },
    { from: 'marcus@unit12.com', subject: 'Heater not working' },
    { from: 'jen@unit7.com', subject: 'Light bulb replacement' },
  ]
  return (
    <div className="rounded-xl border border-gray-200 bg-white p-4 space-y-2">
      <div className="flex items-center gap-1.5 text-[11px] font-semibold text-gray-700 mb-1">
        <Inbox className="w-3.5 h-3.5 text-indigo-600" /> Incoming requests
      </div>
      {emails.map((e, i) => (
        <div key={i} className={`flex items-center gap-3 rounded-lg border border-gray-100 bg-white px-3 py-2.5 ${enter}`} style={{ animationDelay: `${i * 180}ms` }}>
          <div className="w-7 h-7 rounded-full bg-indigo-100 flex items-center justify-center shrink-0">
            <span className="text-[10px] font-bold text-indigo-600">{e.from[0].toUpperCase()}</span>
          </div>
          <div className="flex-1 min-w-0">
            <div className="text-xs font-medium text-gray-800 truncate">{e.from}</div>
            <div className="text-[11px] text-gray-500 truncate">{e.subject}</div>
          </div>
          <ArrowRight className="w-3 h-3 text-gray-300 shrink-0" />
          <span className="text-[9px] font-semibold bg-green-50 text-green-700 px-1.5 py-0.5 rounded-full shrink-0">Ticket</span>
        </div>
      ))}
      <div className="flex items-center gap-1.5 pt-1">
        <span className="w-2 h-2 rounded-full bg-green-400 animate-pulse" />
        <span className="text-[10px] text-gray-400">Auto-captured from email, web form &amp; SMS</span>
      </div>
    </div>
  )
}

function DispatchVisual() {
  const vendors = [
    { name: 'ProFlow Plumbing', score: 96, rating: '4.9', selected: true },
    { name: 'AquaFix Co.', score: 82, rating: '4.7', selected: false },
    { name: 'PipeMasters', score: 68, rating: '4.5', selected: false },
  ]
  return (
    <div className="rounded-xl border border-gray-200 bg-white p-4 space-y-2.5">
      <div className="flex items-center gap-1.5 text-[11px] font-semibold text-gray-700 mb-1">
        <Send className="w-3.5 h-3.5 text-blue-600" /> AI vendor match — Plumbing
      </div>
      {vendors.map((v, i) => (
        <div key={i} className={`rounded-lg border px-3 py-2 ${v.selected ? 'border-blue-300 ring-1 ring-blue-200 bg-blue-50/40' : 'border-gray-100'} ${enter}`} style={{ animationDelay: `${i * 150}ms` }}>
          <div className="flex items-center gap-2 mb-1.5">
            <div className="w-6 h-6 rounded-full bg-blue-100 flex items-center justify-center shrink-0">
              <span className="text-[10px] font-bold text-blue-600">{v.name[0]}</span>
            </div>
            <span className="text-[11px] font-medium text-gray-800 flex-1">{v.name}</span>
            <span className="text-[10px] text-gray-400">⭐ {v.rating}</span>
            {v.selected && <span className="text-[9px] font-semibold bg-blue-600 text-white px-2 py-0.5 rounded-full">Best match</span>}
          </div>
          <div className="flex items-center gap-2">
            <div className="flex-1 h-1.5 rounded-full bg-gray-100 overflow-hidden">
              <div className="h-full rounded-full bg-gradient-to-r from-blue-500 to-cyan-500 transition-all duration-700" style={{ width: `${v.score}%` }} />
            </div>
            <span className="text-[10px] font-semibold text-gray-500 w-8 text-right">{v.score}%</span>
          </div>
        </div>
      ))}
      <p className="text-[10px] text-gray-400 pt-0.5">Ranked by trade, location, availability &amp; past performance.</p>
    </div>
  )
}

function TrackVisual() {
  const cols = [
    { label: 'New', tone: 'bg-gray-100 text-gray-600', cards: ['Door lock'] },
    { label: 'Assigned', tone: 'bg-blue-50 text-blue-700', cards: ['Kitchen leak', 'Paint'] },
    { label: 'In Progress', tone: 'bg-yellow-50 text-yellow-700', cards: ['Heater fix'] },
    { label: 'Done', tone: 'bg-green-50 text-green-700', cards: ['Bulb', 'Window'] },
  ]
  return (
    <div className="rounded-xl border border-gray-200 bg-white p-4">
      <div className="flex items-center gap-1.5 text-[11px] font-semibold text-gray-700 mb-3">
        <ClipboardList className="w-3.5 h-3.5 text-emerald-600" /> Live ticket board
      </div>
      <div className="grid grid-cols-4 gap-2">
        {cols.map((c, ci) => (
          <div key={c.label}>
            <div className={`flex items-center gap-1 rounded-md px-1.5 py-0.5 mb-2 ${c.tone}`}>
              <span className="text-[9px] font-semibold">{c.label}</span>
            </div>
            <div className="space-y-1.5">
              {c.cards.map((card, idx) => (
                <div key={card} className={`rounded border border-gray-100 bg-white p-1.5 ${enter}`} style={{ animationDelay: `${(ci * 2 + idx) * 90}ms` }}>
                  <div className="h-1.5 bg-gray-200 rounded w-full mb-1" />
                  <div className="text-[8px] text-gray-400 truncate">{card}</div>
                </div>
              ))}
            </div>
          </div>
        ))}
      </div>
      <p className="text-[10px] text-gray-400 pt-3">Drag-and-drop or auto-advance — status syncs everywhere instantly.</p>
    </div>
  )
}

function CommunicateVisual() {
  const events = [
    { who: 'Tenant', msg: 'Request received — we are on it', time: 'Now', tone: 'bg-indigo-50 text-indigo-700' },
    { who: 'Vendor', msg: 'ProFlow Plumbing assigned & notified', time: '+2 min', tone: 'bg-blue-50 text-blue-700' },
    { who: 'Tenant', msg: 'Plumber arriving 2–4pm today', time: '+1 hr', tone: 'bg-emerald-50 text-emerald-700' },
    { who: 'Tenant', msg: 'Job complete — how did we do?', time: 'Done', tone: 'bg-amber-50 text-amber-700' },
  ]
  return (
    <div className="rounded-xl border border-gray-200 bg-white p-4">
      <div className="flex items-center gap-1.5 text-[11px] font-semibold text-gray-700 mb-3">
        <Bell className="w-3.5 h-3.5 text-amber-600" /> Automated updates
      </div>
      <div className="space-y-2.5 relative">
        <div className="absolute left-[7px] top-1 bottom-1 w-px bg-gray-100" />
        {events.map((e, i) => (
          <div key={i} className={`flex items-start gap-2.5 relative ${enter}`} style={{ animationDelay: `${i * 160}ms` }}>
            <span className="w-3.5 h-3.5 rounded-full bg-white border-2 border-amber-400 shrink-0 mt-0.5 z-10" />
            <div className="flex-1 rounded-lg border border-gray-100 px-2.5 py-1.5">
              <div className="flex items-center justify-between gap-2">
                <span className={`text-[9px] font-semibold px-1.5 py-0.5 rounded-full ${e.tone}`}>{e.who}</span>
                <span className="text-[9px] text-gray-400">{e.time}</span>
              </div>
              <p className="text-[11px] text-gray-600 mt-1">{e.msg}</p>
            </div>
          </div>
        ))}
      </div>
      <p className="text-[10px] text-gray-400 pt-3">Zero manual follow-ups — everyone stays informed automatically.</p>
    </div>
  )
}

function ResolveVisual() {
  const stats = [
    { label: 'Resolved', value: '34', tone: 'text-emerald-600' },
    { label: 'Avg. time', value: '2.1d', tone: 'text-blue-600' },
    { label: 'SLA hit', value: '97%', tone: 'text-indigo-600' },
  ]
  const vendors = [
    { name: 'ProFlow Plumbing', stars: 5, jobs: 12 },
    { name: 'PowerUp Electric', stars: 5, jobs: 14 },
    { name: 'AquaFix Co.', stars: 4, jobs: 8 },
  ]
  return (
    <div className="rounded-xl border border-gray-200 bg-white p-4 space-y-2.5">
      <div className="flex items-center gap-1.5 text-[11px] font-semibold text-gray-700">
        <BarChart3 className="w-3.5 h-3.5 text-rose-600" /> Monthly performance
      </div>
      <div className="grid grid-cols-3 gap-2">
        {stats.map((s, i) => (
          <div key={s.label} className={`rounded-lg border border-gray-100 bg-gray-50 py-2 text-center ${enter}`} style={{ animationDelay: `${i * 100}ms` }}>
            <div className={`text-lg font-bold ${s.tone}`}>{s.value}</div>
            <div className="text-[9px] text-gray-400">{s.label}</div>
          </div>
        ))}
      </div>
      <div className="space-y-1.5">
        {vendors.map((v, i) => (
          <div key={v.name} className={`flex items-center gap-2 rounded-lg border border-gray-100 px-3 py-1.5 ${enter}`} style={{ animationDelay: `${(i + 3) * 100}ms` }}>
            <CheckCircle2 className="w-3 h-3 text-emerald-500 shrink-0" />
            <span className="text-[10px] text-gray-700 flex-1">{v.name}</span>
            <span className="text-[10px] text-yellow-500">{'★'.repeat(v.stars)}</span>
            <span className="text-[10px] text-gray-400">{v.jobs} jobs</span>
          </div>
        ))}
      </div>
      <div className="flex items-center gap-1.5 text-[10px] text-gray-400 pt-0.5">
        <ShieldCheck className="w-3 h-3" /> Full audit trail logged for every action.
      </div>
    </div>
  )
}

/* ============================ STEP DATA ============================ */
type Step = {
  label: string; tagline: string; title: string; description: string
  benefits: string[]; icon: typeof Zap; accent: Accent; render: (a: Accent) => React.ReactNode
}
const STEPS: Step[] = [
  {
    label: 'Capture', tagline: 'Step 1', title: 'Every request, captured automatically',
    description: 'Tenants reach out however they like — email, web form, or SMS. Maintena instantly turns each message into a structured, trackable ticket. Nothing slips through the cracks.',
    benefits: ['No portals or apps for tenants to learn', 'Works from a forwarded email address', 'Photos & details attached automatically'],
    icon: Inbox, accent: ACCENTS.indigo, render: () => <CaptureVisual />,
  },
  {
    label: 'AI Triage', tagline: 'Step 2', title: 'AI reads & understands every request',
    description: 'Our AI instantly classifies the trade, scores urgency, flags safety risks, and writes a clean summary — work that used to take a coordinator minutes per ticket. Try it yourself below.',
    benefits: ['Catches emergencies before they escalate', 'Consistent prioritization, 24/7', 'Plain-English summary for your team'],
    icon: Sparkles, accent: ACCENTS.violet, render: (a) => <AiTriageDemo accent={a} />,
  },
  {
    label: 'Dispatch', tagline: 'Step 3', title: 'The right vendor, matched instantly',
    description: 'Maintena ranks your vendors by trade, location, availability, and track record — then dispatches with one click. The vendor gets everything they need automatically.',
    benefits: ['One-click assign & notify', 'Balances cost, speed & quality', 'Keeps your best vendors busy'],
    icon: Users, accent: ACCENTS.blue, render: () => <DispatchVisual />,
  },
  {
    label: 'Track', tagline: 'Step 4', title: 'See everything on one live board',
    description: 'A real-time Kanban shows every ticket status at a glance. SLA timers warn you before anything goes overdue, so nothing stalls.',
    benefits: ['Drag-and-drop or auto-advance', 'Filter by property, urgency or vendor', 'SLA alerts before deadlines slip'],
    icon: ClipboardList, accent: ACCENTS.emerald, render: () => <TrackVisual />,
  },
  {
    label: 'Communicate', tagline: 'Step 5', title: 'Everyone stays in the loop — automatically',
    description: 'Tenants and vendors get timely updates at every stage without you lifting a finger. Fewer status calls, happier residents.',
    benefits: ['Automatic updates at each milestone', 'Branded, professional messaging', 'Cuts inbound “any update?” calls'],
    icon: Bell, accent: ACCENTS.amber, render: () => <CommunicateVisual />,
  },
  {
    label: 'Resolve', tagline: 'Step 6', title: 'Close the loop with full accountability',
    description: 'Mark complete, log costs, and collect ratings. Maintena reports resolution times, SLA performance, and vendor scorecards — with a complete audit trail.',
    benefits: ['Vendor scorecards & ratings', 'SLA & cost reporting', 'Timestamped audit trail for every action'],
    icon: BarChart3, accent: ACCENTS.rose, render: () => <ResolveVisual />,
  },
]
const INTERACTIVE_STEP = 1
const AUTOPLAY_MS = 7000

/* ============================ MODAL ============================ */
export function HowItWorksModal({ open, onClose }: { open: boolean; onClose: () => void }) {
  const [active, setActive] = useState(0)
  const [playing, setPlaying] = useState(true)

  useEffect(() => {
    if (open) { setActive(0); setPlaying(true) }
  }, [open])

  useEffect(() => {
    if (!open) return
    const goRelative = (delta: number) => { setPlaying(false); setActive(s => Math.max(0, Math.min(STEPS.length - 1, s + delta))) }
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
      if (e.key === 'ArrowRight') goRelative(1)
      if (e.key === 'ArrowLeft') goRelative(-1)
    }
    window.addEventListener('keydown', onKey)
    document.body.style.overflow = 'hidden'
    return () => { window.removeEventListener('keydown', onKey); document.body.style.overflow = '' }
  }, [open, onClose])

  useEffect(() => {
    if (!open || !playing) return
    if (active === INTERACTIVE_STEP) return
    const id = setTimeout(() => {
      setActive(s => {
        if (s >= STEPS.length - 1) { setPlaying(false); return s }
        return s + 1
      })
    }, AUTOPLAY_MS)
    return () => clearTimeout(id)
  }, [open, playing, active])

  const goTo = (i: number) => { setPlaying(false); setActive(i) }

  if (!open) return null
  const step = STEPS[active]
  const Icon = step.icon
  const a = step.accent
  const progress = ((active + 1) / STEPS.length) * 100

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-4" onClick={onClose}>
      <div className="absolute inset-0 bg-gray-900/60 backdrop-blur-sm animate-in fade-in" />
      <div
        className="relative w-full max-w-4xl max-h-[92vh] bg-white rounded-2xl shadow-2xl overflow-hidden flex flex-col animate-in fade-in zoom-in-95 duration-200"
        onClick={e => e.stopPropagation()}
      >
        {/* progress bar */}
        <div className="h-1 w-full bg-gray-100">
          <div className={`h-full bg-gradient-to-r ${a.bar} transition-all duration-500`} style={{ width: `${progress}%` }} />
        </div>

        {/* header */}
        <div className="flex items-center justify-between px-5 sm:px-6 py-3.5 border-b border-gray-100">
          <div className="flex items-center gap-2.5">
            <div className="w-7 h-7 rounded-lg bg-indigo-600 flex items-center justify-center">
              <MaintenaMark className="w-4 h-4 text-white" />
            </div>
            <div>
              <h2 className="text-sm font-semibold text-gray-900 leading-none">How Maintena works</h2>
              <p className="text-[11px] text-gray-400 mt-1">From request to resolution — in minutes</p>
            </div>
          </div>
          <div className="flex items-center gap-1">
            <button
              onClick={() => setPlaying(p => !p)}
              className="w-8 h-8 rounded-lg flex items-center justify-center text-gray-400 hover:bg-gray-100 hover:text-gray-700 transition-colors"
              title={playing ? 'Pause tour' : 'Play tour'}
            >
              {playing ? <Pause className="w-4 h-4" /> : <Play className="w-4 h-4" />}
            </button>
            <button
              onClick={onClose}
              className="w-8 h-8 rounded-lg flex items-center justify-center text-gray-400 hover:bg-gray-100 hover:text-gray-700 transition-colors"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        </div>

        {/* body: rail + content */}
        <div className="flex flex-1 min-h-0">
          {/* left rail (desktop) */}
          <nav className="hidden md:flex flex-col gap-1 w-52 shrink-0 border-r border-gray-100 p-3 overflow-y-auto">
            {STEPS.map((s, i) => {
              const done = i < active
              const cur = i === active
              const sa = s.accent
              return (
                <button
                  key={i}
                  onClick={() => goTo(i)}
                  className={`flex items-center gap-2.5 rounded-lg px-2.5 py-2 text-left transition-colors ${cur ? 'bg-gray-50' : 'hover:bg-gray-50'}`}
                >
                  <span className={`w-6 h-6 rounded-full flex items-center justify-center text-[10px] font-bold shrink-0 transition-colors ${cur ? `${sa.dot} text-white` : done ? 'bg-emerald-100 text-emerald-600' : 'bg-gray-100 text-gray-400'}`}>
                    {done ? <CheckCircle2 className="w-3.5 h-3.5" /> : i + 1}
                  </span>
                  <span className={`text-xs font-medium ${cur ? 'text-gray-900' : 'text-gray-500'}`}>{s.label}</span>
                </button>
              )
            })}
          </nav>

          {/* content */}
          <div className="flex-1 min-w-0 overflow-y-auto">
            {/* mobile pills */}
            <div className="md:hidden flex gap-1.5 px-4 pt-3 overflow-x-auto">
              {STEPS.map((s, i) => (
                <button
                  key={i}
                  onClick={() => goTo(i)}
                  className={`text-[11px] font-medium px-2.5 py-1 rounded-full whitespace-nowrap transition-colors ${i === active ? `${s.accent.dot} text-white` : 'bg-gray-100 text-gray-500'}`}
                >
                  {i + 1}. {s.label}
                </button>
              ))}
            </div>

            <div key={active} className="p-5 sm:p-6 grid md:grid-cols-2 gap-5 items-start animate-in fade-in slide-in-from-right-3 duration-300">
              {/* left: copy */}
              <div>
                <div className={`inline-flex items-center gap-1.5 text-[11px] font-semibold ${a.pill} ${a.pillText} px-2.5 py-1 rounded-full mb-3`}>
                  <Icon className="w-3.5 h-3.5" /> {step.tagline}
                </div>
                <h3 className="text-lg sm:text-xl font-semibold text-gray-900 leading-snug mb-2">{step.title}</h3>
                <p className="text-sm text-gray-500 leading-relaxed mb-4">{step.description}</p>
                <ul className="space-y-2">
                  {step.benefits.map((b, i) => (
                    <li key={b} className={`flex items-start gap-2 text-[13px] text-gray-600 ${enter}`} style={{ animationDelay: `${i * 80}ms` }}>
                      <CheckCircle2 className={`w-4 h-4 mt-0.5 shrink-0 ${a.iconText}`} />
                      {b}
                    </li>
                  ))}
                </ul>
              </div>
              {/* right: visual */}
              <div>{step.render(a)}</div>
            </div>
          </div>
        </div>

        {/* footer */}
        <div className="flex items-center justify-between px-5 sm:px-6 py-3.5 border-t border-gray-100">
          <button
            onClick={() => goTo(Math.max(0, active - 1))}
            disabled={active === 0}
            className="inline-flex items-center gap-1 text-sm text-gray-500 hover:text-gray-800 disabled:opacity-30 disabled:cursor-not-allowed transition-colors"
          >
            <ArrowLeft className="w-3.5 h-3.5" /> Back
          </button>

          <div className="flex items-center gap-1.5">
            {STEPS.map((_, i) => (
              <button
                key={i}
                onClick={() => goTo(i)}
                className={`h-1.5 rounded-full transition-all ${i === active ? `w-5 ${a.dot}` : 'w-1.5 bg-gray-200 hover:bg-gray-300'}`}
                aria-label={`Go to step ${i + 1}`}
              />
            ))}
          </div>

          {active < STEPS.length - 1 ? (
            <button
              onClick={() => goTo(active + 1)}
              className="inline-flex items-center gap-1 text-sm font-medium text-indigo-600 hover:text-indigo-700 transition-colors"
            >
              Next <ArrowRight className="w-3.5 h-3.5" />
            </button>
          ) : (
            <a
              href="/signup"
              className="inline-flex items-center gap-1.5 text-sm font-medium bg-indigo-600 text-white px-4 py-1.5 rounded-lg hover:bg-indigo-700 transition-colors"
            >
              Start free <ArrowRight className="w-3.5 h-3.5" />
            </a>
          )}
        </div>
      </div>
    </div>
  )
}

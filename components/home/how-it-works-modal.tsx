'use client'

import { useState, useEffect } from 'react'
import { X, ArrowRight, Zap, Users, ClipboardList, Bell, BarChart3, CheckCircle2 } from 'lucide-react'

const steps = [
  {
    step: '01',
    title: 'Tenant submits a request',
    description: 'Tenants email or submit a maintenance request. Maintena instantly captures it and creates a structured ticket — no forms, no portals.',
    icon: Bell,
    color: 'indigo',
    visual: <InboxVisual />,
  },
  {
    step: '02',
    title: 'AI classifies & prioritizes',
    description: 'Our AI reads the request, determines the trade type (plumbing, electrical, HVAC…), sets urgency, and suggests the right vendor — in seconds.',
    icon: Zap,
    color: 'violet',
    visual: <ClassifyVisual />,
  },
  {
    step: '03',
    title: 'Vendor is dispatched',
    description: 'You approve with one click. The vendor gets notified automatically with all the details they need. No calls, no back-and-forth.',
    icon: Users,
    color: 'blue',
    visual: <DispatchVisual />,
  },
  {
    step: '04',
    title: 'Track in real time',
    description: 'Every ticket lives on your Kanban board. Automated updates go to tenants at each stage. SLA timers alert you before anything goes overdue.',
    icon: ClipboardList,
    color: 'emerald',
    visual: <TrackVisual />,
  },
  {
    step: '05',
    title: 'Resolve & report',
    description: 'Mark complete, log costs, and review performance. Full audit trail, resolution times, and vendor ratings — all in one place.',
    icon: BarChart3,
    color: 'orange',
    visual: <ResolveVisual />,
  },
]

function InboxVisual() {
  return (
    <div className="bg-gray-50 rounded-xl border border-gray-200 p-4 space-y-2">
      {[
        { from: 'sarah@apt4b.com', subject: 'Leak under kitchen sink', time: '9:41 AM', urgent: true },
        { from: 'marcus@unit12.com', subject: 'Heater not working', time: '8:15 AM', urgent: true },
        { from: 'jen@unit7.com', subject: 'Light bulb replacement needed', time: 'Yesterday', urgent: false },
      ].map((email, i) => (
        <div key={i} className="bg-white rounded-lg border border-gray-100 px-3 py-2.5 flex items-start gap-3">
          <div className="w-7 h-7 rounded-full bg-indigo-100 flex items-center justify-center shrink-0 mt-0.5">
            <span className="text-[10px] font-bold text-indigo-600">{email.from[0].toUpperCase()}</span>
          </div>
          <div className="flex-1 min-w-0">
            <div className="flex items-center justify-between gap-2">
              <span className="text-xs font-medium text-gray-800 truncate">{email.from}</span>
              <span className="text-[10px] text-gray-400 shrink-0">{email.time}</span>
            </div>
            <p className="text-[11px] text-gray-500 truncate">{email.subject}</p>
          </div>
          {email.urgent && <span className="text-[9px] font-semibold bg-red-50 text-red-600 px-1.5 py-0.5 rounded-full shrink-0">Urgent</span>}
        </div>
      ))}
      <div className="flex items-center gap-1.5 pt-1">
        <div className="w-2 h-2 rounded-full bg-green-400 animate-pulse" />
        <span className="text-[10px] text-gray-400">3 new requests captured automatically</span>
      </div>
    </div>
  )
}

function ClassifyVisual() {
  return (
    <div className="bg-gray-50 rounded-xl border border-gray-200 p-4">
      <div className="bg-white rounded-lg border border-gray-100 p-3 mb-3">
        <p className="text-[11px] text-gray-500 italic">&ldquo;There&apos;s a leak under my kitchen sink and water is getting on the floor&rdquo;</p>
      </div>
      <div className="flex items-center gap-2 mb-3">
        <div className="h-px flex-1 bg-gray-200" />
        <div className="flex items-center gap-1 text-[10px] text-violet-600 font-medium">
          <Zap className="w-3 h-3" /> AI Processing
        </div>
        <div className="h-px flex-1 bg-gray-200" />
      </div>
      <div className="grid grid-cols-2 gap-2">
        {[
          { label: 'Trade', value: 'Plumbing' },
          { label: 'Urgency', value: '🔴 High' },
          { label: 'Category', value: 'Water Damage Risk' },
          { label: 'ETA', value: 'Same day' },
        ].map(item => (
          <div key={item.label} className="bg-violet-50 rounded-md px-2.5 py-1.5">
            <div className="text-[9px] font-medium text-violet-400 uppercase tracking-wide">{item.label}</div>
            <div className="text-[11px] font-semibold text-violet-800">{item.value}</div>
          </div>
        ))}
      </div>
    </div>
  )
}

function DispatchVisual() {
  return (
    <div className="bg-gray-50 rounded-xl border border-gray-200 p-4 space-y-2">
      <div className="text-[10px] font-semibold text-gray-500 uppercase tracking-wide mb-1">Recommended vendors</div>
      {[
        { name: 'ProFlow Plumbing', rating: '4.9', jobs: '42 jobs', available: true, selected: true },
        { name: 'AquaFix Co.', rating: '4.7', jobs: '28 jobs', available: true, selected: false },
        { name: 'PipeMasters', rating: '4.5', jobs: '15 jobs', available: false, selected: false },
      ].map((v, i) => (
        <div key={i} className={`bg-white rounded-lg border px-3 py-2 flex items-center gap-3 ${v.selected ? 'border-blue-300 ring-1 ring-blue-200' : 'border-gray-100'}`}>
          <div className="w-7 h-7 rounded-full bg-blue-100 flex items-center justify-center shrink-0">
            <span className="text-[10px] font-bold text-blue-600">{v.name[0]}</span>
          </div>
          <div className="flex-1">
            <div className="text-[11px] font-medium text-gray-800">{v.name}</div>
            <div className="text-[10px] text-gray-400">⭐ {v.rating} · {v.jobs}</div>
          </div>
          {v.selected
            ? <span className="text-[9px] font-semibold bg-blue-600 text-white px-2 py-0.5 rounded-full">Assign</span>
            : <span className={`text-[9px] font-medium ${v.available ? 'text-green-600' : 'text-gray-400'}`}>{v.available ? 'Available' : 'Busy'}</span>
          }
        </div>
      ))}
    </div>
  )
}

function TrackVisual() {
  return (
    <div className="bg-gray-50 rounded-xl border border-gray-200 p-4">
      <div className="grid grid-cols-4 gap-2">
        {[
          { label: 'New', count: 2, color: 'bg-gray-100', textColor: 'text-gray-600', cards: ['HVAC noise', 'Door lock'] },
          { label: 'Assigned', count: 3, color: 'bg-blue-50', textColor: 'text-blue-700', cards: ['Kitchen leak', 'Paint'] },
          { label: 'In Progress', count: 2, color: 'bg-yellow-50', textColor: 'text-yellow-700', cards: ['Heater fix'] },
          { label: 'Done', count: 8, color: 'bg-green-50', textColor: 'text-green-700', cards: ['Bulb', 'Window'] },
        ].map(col => (
          <div key={col.label}>
            <div className={`flex items-center gap-1 rounded-md px-1.5 py-0.5 mb-2 ${col.color}`}>
              <span className={`text-[9px] font-semibold ${col.textColor}`}>{col.label}</span>
              <span className={`text-[9px] ${col.textColor} opacity-70`}>{col.count}</span>
            </div>
            <div className="space-y-1.5">
              {col.cards.map(card => (
                <div key={card} className="bg-white rounded border border-gray-100 p-1.5">
                  <div className="h-1.5 bg-gray-200 rounded w-full mb-1" />
                  <div className="text-[8px] text-gray-400 truncate">{card}</div>
                </div>
              ))}
            </div>
          </div>
        ))}
      </div>
    </div>
  )
}

function ResolveVisual() {
  return (
    <div className="bg-gray-50 rounded-xl border border-gray-200 p-4 space-y-2">
      <div className="bg-white rounded-lg border border-gray-100 p-3">
        <div className="flex items-center justify-between mb-2">
          <span className="text-[11px] font-semibold text-gray-800">Monthly Summary</span>
          <span className="text-[10px] text-gray-400">May 2025</span>
        </div>
        <div className="grid grid-cols-3 gap-2">
          {[
            { label: 'Resolved', value: '34', color: 'text-green-600' },
            { label: 'Avg. time', value: '2.1d', color: 'text-blue-600' },
            { label: 'SLA hit', value: '97%', color: 'text-indigo-600' },
          ].map(stat => (
            <div key={stat.label} className="text-center">
              <div className={`text-base font-bold ${stat.color}`}>{stat.value}</div>
              <div className="text-[9px] text-gray-400">{stat.label}</div>
            </div>
          ))}
        </div>
      </div>
      <div className="space-y-1.5">
        {[
          { label: 'ProFlow Plumbing', rating: 5, jobs: 12 },
          { label: 'AquaFix Co.', rating: 4, jobs: 8 },
          { label: 'PowerUp Electric', rating: 5, jobs: 14 },
        ].map(v => (
          <div key={v.label} className="bg-white rounded-lg border border-gray-100 px-3 py-1.5 flex items-center gap-2">
            <CheckCircle2 className="w-3 h-3 text-green-500 shrink-0" />
            <span className="text-[10px] text-gray-700 flex-1">{v.label}</span>
            <span className="text-[10px] text-yellow-500">{'★'.repeat(v.rating)}</span>
            <span className="text-[10px] text-gray-400">{v.jobs} jobs</span>
          </div>
        ))}
      </div>
    </div>
  )
}

export function HowItWorksModal({ open, onClose }: { open: boolean; onClose: () => void }) {
  const [activeStep, setActiveStep] = useState(0)

  useEffect(() => {
    if (!open) return
    const handler = (e: KeyboardEvent) => { if (e.key === 'Escape') onClose() }
    window.addEventListener('keydown', handler)
    return () => window.removeEventListener('keydown', handler)
  }, [open, onClose])

  useEffect(() => {
    if (open) document.body.style.overflow = 'hidden'
    else document.body.style.overflow = ''
    return () => { document.body.style.overflow = '' }
  }, [open])

  if (!open) return null

  const step = steps[activeStep]
  const Icon = step.icon

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4" onClick={onClose}>
      <div className="absolute inset-0 bg-black/50 backdrop-blur-sm" />
      <div
        className="relative bg-white rounded-2xl shadow-2xl w-full max-w-2xl max-h-[90vh] overflow-hidden flex flex-col"
        onClick={e => e.stopPropagation()}
      >
        {/* Header */}
        <div className="flex items-center justify-between px-6 py-4 border-b border-gray-100">
          <div>
            <h2 className="text-lg font-semibold text-gray-900">How Maintena works</h2>
            <p className="text-sm text-gray-500">From request to resolution in minutes</p>
          </div>
          <button
            onClick={onClose}
            className="w-8 h-8 rounded-lg flex items-center justify-center text-gray-400 hover:bg-gray-100 hover:text-gray-600 transition-colors"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Step tabs */}
        <div className="flex gap-1 px-6 pt-4 overflow-x-auto">
          {steps.map((s, i) => (
            <button
              key={i}
              onClick={() => setActiveStep(i)}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium whitespace-nowrap transition-colors ${
                activeStep === i
                  ? 'bg-indigo-600 text-white'
                  : 'bg-gray-100 text-gray-600 hover:bg-gray-200'
              }`}
            >
              <span className="font-mono text-[10px] opacity-70">{s.step}</span>
              {s.title.split(' ').slice(0, 2).join(' ')}
            </button>
          ))}
        </div>

        {/* Content */}
        <div className="flex-1 overflow-y-auto p-6">
          <div className="flex items-start gap-3 mb-5">
            <div className={`w-10 h-10 rounded-xl flex items-center justify-center shrink-0 bg-indigo-50`}>
              <Icon className="w-5 h-5 text-indigo-600" />
            </div>
            <div>
              <div className="text-xs font-mono text-gray-400 mb-0.5">Step {step.step}</div>
              <h3 className="text-base font-semibold text-gray-900">{step.title}</h3>
              <p className="text-sm text-gray-500 mt-1 leading-relaxed">{step.description}</p>
            </div>
          </div>

          {step.visual}
        </div>

        {/* Footer nav */}
        <div className="px-6 py-4 border-t border-gray-100 flex items-center justify-between">
          <button
            onClick={() => setActiveStep(i => Math.max(0, i - 1))}
            disabled={activeStep === 0}
            className="text-sm text-gray-500 hover:text-gray-800 disabled:opacity-30 disabled:cursor-not-allowed transition-colors"
          >
            ← Previous
          </button>
          <div className="flex gap-1.5">
            {steps.map((_, i) => (
              <button
                key={i}
                onClick={() => setActiveStep(i)}
                className={`w-1.5 h-1.5 rounded-full transition-colors ${i === activeStep ? 'bg-indigo-600' : 'bg-gray-200'}`}
              />
            ))}
          </div>
          {activeStep < steps.length - 1 ? (
            <button
              onClick={() => setActiveStep(i => Math.min(steps.length - 1, i + 1))}
              className="flex items-center gap-1 text-sm font-medium text-indigo-600 hover:text-indigo-700 transition-colors"
            >
              Next <ArrowRight className="w-3 h-3" />
            </button>
          ) : (
            <a
              href="/login"
              className="flex items-center gap-1 text-sm font-medium bg-indigo-600 text-white px-4 py-1.5 rounded-lg hover:bg-indigo-700 transition-colors"
            >
              Get started <ArrowRight className="w-3 h-3" />
            </a>
          )}
        </div>
      </div>
    </div>
  )
}

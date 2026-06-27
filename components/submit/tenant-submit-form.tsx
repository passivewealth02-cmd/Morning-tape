'use client'

import { useState, useMemo } from 'react'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Textarea } from '@/components/ui/textarea'
import { CheckCircle2, MapPin, ArrowRight } from 'lucide-react'

type PropertyOption = { id: string; name: string; address: string }
type UnitOption = { id: string; property_id: string; unit_number: string }

const URGENCY_OPTIONS = [
  { value: 'low', label: 'Low', hint: 'Can wait a few days', color: 'border-gray-200 bg-white', active: 'border-gray-400 bg-gray-50' },
  { value: 'medium', label: 'Medium', hint: 'Needs attention soon', color: 'border-gray-200 bg-white', active: 'border-indigo-400 bg-indigo-50' },
  { value: 'high', label: 'High', hint: 'Significant problem', color: 'border-gray-200 bg-white', active: 'border-orange-400 bg-orange-50' },
  { value: 'emergency', label: 'Emergency', hint: 'Safety risk / major damage', color: 'border-gray-200 bg-white', active: 'border-red-400 bg-red-50' },
]

export function TenantSubmitForm({
  token,
  properties,
  units,
  presetProperty,
  presetUnit,
}: {
  token: string
  properties: PropertyOption[]
  units: UnitOption[]
  presetProperty: string
  presetUnit: string
}) {
  const hasProperties = properties.length > 0
  const locked = Boolean(presetProperty) // came from a QR / deep link

  const [form, setForm] = useState({
    title: '',
    description: '',
    urgency: 'medium',
    property_id: presetProperty || '',
    unit_id: presetUnit || '',
    property_hint: '',
    unit_number: '',
    tenant_name: '',
    tenant_email: '',
    tenant_phone: '',
  })
  // If no properties exist for this org, fall straight to free-text entry.
  const [notListed, setNotListed] = useState(!hasProperties)
  const [status, setStatus] = useState<'idle' | 'loading' | 'success' | 'error'>('idle')
  const [error, setError] = useState('')
  const [ticketId, setTicketId] = useState<string | null>(null)

  const set = (key: string, value: string) => setForm(f => ({ ...f, [key]: value }))

  const unitsForProperty = useMemo(
    () => units.filter(u => u.property_id === form.property_id),
    [units, form.property_id]
  )

  const presetLabel = useMemo(() => {
    if (!locked) return ''
    const p = properties.find(x => x.id === presetProperty)
    const u = units.find(x => x.id === presetUnit)
    if (!p) return ''
    return u ? `${p.name} · Unit ${u.unit_number}` : p.name
  }, [locked, presetProperty, presetUnit, properties, units])

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    setStatus('loading')
    setError('')
    try {
      const payload: Record<string, string> = {
        title: form.title,
        description: form.description,
        urgency: form.urgency,
        tenant_name: form.tenant_name,
        tenant_email: form.tenant_email,
        tenant_phone: form.tenant_phone,
      }
      if (notListed) {
        payload.property_hint = form.property_hint
        payload.unit_number = form.unit_number
      } else {
        payload.property_id = form.property_id
        payload.unit_id = form.unit_id
      }

      const res = await fetch(`/api/submit/${token}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      })
      const data = await res.json()
      if (!res.ok) throw new Error(data.error || 'Failed to submit request')
      setTicketId(data.ticket_id ?? null)
      setStatus('success')
    } catch (err) {
      setStatus('error')
      setError(err instanceof Error ? err.message : 'An error occurred')
    }
  }

  if (status === 'success') {
    return (
      <div className="bg-white rounded-2xl border border-gray-200 p-8 text-center shadow-sm">
        <div className="w-14 h-14 rounded-full bg-green-50 flex items-center justify-center mx-auto mb-4">
          <CheckCircle2 className="w-7 h-7 text-green-600" />
        </div>
        <h2 className="text-lg font-semibold text-gray-900">Request submitted</h2>
        <p className="text-sm text-gray-500 mt-2">
          Thanks! Your property manager has been notified and the right person will be assigned.
          {form.tenant_email && ' You’ll get email updates as it progresses.'}
        </p>
        {ticketId && (
          <a
            href={`/track/${ticketId}`}
            className="mt-5 inline-flex items-center gap-1.5 text-sm font-semibold text-indigo-600 hover:text-indigo-700"
          >
            Track your request
            <ArrowRight className="w-4 h-4" />
          </a>
        )}
        <div className="mt-5">
          <button
            onClick={() => {
              setForm(f => ({
                ...f,
                title: '', description: '', urgency: 'medium',
              }))
              setTicketId(null)
              setStatus('idle')
            }}
            className="text-sm font-medium text-gray-500 hover:text-gray-700"
          >
            Submit another request
          </button>
        </div>
      </div>
    )
  }

  const loading = status === 'loading'

  return (
    <form onSubmit={handleSubmit} className="space-y-4">
      {/* Location */}
      <div className="bg-white rounded-2xl border border-gray-200 p-5 shadow-sm space-y-4">
        {locked && presetLabel ? (
          <div className="flex items-center gap-2.5 rounded-xl bg-indigo-50 border border-indigo-100 px-4 py-3">
            <MapPin className="w-4 h-4 text-indigo-600 shrink-0" />
            <div className="min-w-0">
              <p className="text-xs text-indigo-500 font-medium">Reporting for</p>
              <p className="text-sm font-semibold text-gray-900 truncate">{presetLabel}</p>
            </div>
          </div>
        ) : (
          <>
            <h2 className="text-sm font-semibold text-gray-700">Where is the issue?</h2>
            {hasProperties && !notListed ? (
              <div className="space-y-3">
                <div>
                  <label className="block text-xs font-medium text-gray-600 mb-1.5">Property</label>
                  <select
                    value={form.property_id}
                    onChange={e => setForm(f => ({ ...f, property_id: e.target.value, unit_id: '' }))}
                    disabled={loading}
                    className="w-full rounded-lg border border-gray-200 px-3 py-2.5 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-indigo-200 focus:border-indigo-300"
                  >
                    <option value="">Select your property…</option>
                    {properties.map(p => (
                      <option key={p.id} value={p.id}>{p.name} — {p.address}</option>
                    ))}
                  </select>
                </div>
                {form.property_id && unitsForProperty.length > 0 && (
                  <div>
                    <label className="block text-xs font-medium text-gray-600 mb-1.5">Unit</label>
                    <select
                      value={form.unit_id}
                      onChange={e => set('unit_id', e.target.value)}
                      disabled={loading}
                      className="w-full rounded-lg border border-gray-200 px-3 py-2.5 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-indigo-200 focus:border-indigo-300"
                    >
                      <option value="">Select your unit…</option>
                      {unitsForProperty.map(u => (
                        <option key={u.id} value={u.id}>Unit {u.unit_number}</option>
                      ))}
                    </select>
                  </div>
                )}
                <button
                  type="button"
                  onClick={() => setNotListed(true)}
                  className="text-xs font-medium text-indigo-600 hover:text-indigo-700"
                >
                  My place isn’t listed
                </button>
              </div>
            ) : (
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-medium text-gray-600 mb-1.5">Property / address</label>
                  <Input value={form.property_hint} onChange={e => set('property_hint', e.target.value)} placeholder="120 Maple St" disabled={loading} />
                </div>
                <div>
                  <label className="block text-xs font-medium text-gray-600 mb-1.5">Unit</label>
                  <Input value={form.unit_number} onChange={e => set('unit_number', e.target.value)} placeholder="4B" disabled={loading} />
                </div>
                {hasProperties && (
                  <button
                    type="button"
                    onClick={() => setNotListed(false)}
                    className="col-span-2 text-left text-xs font-medium text-indigo-600 hover:text-indigo-700"
                  >
                    ← Pick from the list instead
                  </button>
                )}
              </div>
            )}
          </>
        )}
      </div>

      {/* Issue */}
      <div className="bg-white rounded-2xl border border-gray-200 p-5 shadow-sm space-y-4">
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1.5">What&apos;s wrong? *</label>
          <Input
            value={form.title}
            onChange={e => set('title', e.target.value)}
            placeholder="e.g. Water leaking under kitchen sink"
            required
            disabled={loading}
          />
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1.5">Add some detail *</label>
          <Textarea
            value={form.description}
            onChange={e => set('description', e.target.value)}
            placeholder="When did it start, how bad is it, and anything else that helps."
            required
            rows={4}
            disabled={loading}
          />
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-2">How urgent is it?</label>
          <div className="grid grid-cols-2 gap-2">
            {URGENCY_OPTIONS.map(opt => {
              const selected = form.urgency === opt.value
              return (
                <button
                  type="button"
                  key={opt.value}
                  onClick={() => set('urgency', opt.value)}
                  disabled={loading}
                  className={`text-left rounded-xl border px-3 py-2.5 transition-colors ${selected ? opt.active : opt.color} ${selected ? '' : 'hover:bg-gray-50'}`}
                >
                  <span className="block text-sm font-semibold text-gray-900">{opt.label}</span>
                  <span className="block text-[11px] text-gray-500 mt-0.5">{opt.hint}</span>
                </button>
              )
            })}
          </div>
        </div>
      </div>

      {/* Contact */}
      <div className="bg-white rounded-2xl border border-gray-200 p-5 shadow-sm space-y-4">
        <div>
          <h2 className="text-sm font-semibold text-gray-700">How can we reach you?</h2>
          <p className="text-xs text-gray-500 mt-0.5">So we can send you updates on this request.</p>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <div>
            <label className="block text-xs font-medium text-gray-600 mb-1.5">Name</label>
            <Input value={form.tenant_name} onChange={e => set('tenant_name', e.target.value)} placeholder="Jane Smith" disabled={loading} />
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-600 mb-1.5">Phone</label>
            <Input value={form.tenant_phone} onChange={e => set('tenant_phone', e.target.value)} placeholder="(555) 000-0000" disabled={loading} />
          </div>
        </div>
        <div>
          <label className="block text-xs font-medium text-gray-600 mb-1.5">Email <span className="text-gray-400">(for updates)</span></label>
          <Input type="email" value={form.tenant_email} onChange={e => set('tenant_email', e.target.value)} placeholder="you@email.com" disabled={loading} />
        </div>
      </div>

      {error && <p className="text-sm text-red-600 text-center">{error}</p>}

      <Button
        type="submit"
        disabled={loading || !form.title || !form.description}
        className="w-full bg-indigo-600 hover:bg-indigo-700 text-white py-3 text-base rounded-xl"
      >
        {loading ? 'Submitting…' : 'Submit request'}
      </Button>
      <p className="text-center text-[11px] text-gray-400">Powered by Maintena · Your manager is notified instantly</p>
    </form>
  )
}

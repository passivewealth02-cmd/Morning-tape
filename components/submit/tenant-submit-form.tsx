'use client'

import { useState } from 'react'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Textarea } from '@/components/ui/textarea'
import { CheckCircle2 } from 'lucide-react'

const URGENCY_OPTIONS = [
  { value: 'low', label: 'Low — can wait a few days' },
  { value: 'medium', label: 'Medium — needs attention soon' },
  { value: 'high', label: 'High — significant problem' },
  { value: 'emergency', label: 'Emergency — safety risk or major damage' },
]

export function TenantSubmitForm({ token }: { token: string }) {
  const [form, setForm] = useState({
    title: '',
    description: '',
    urgency: 'medium',
    property_hint: '',
    unit_number: '',
    tenant_name: '',
    tenant_email: '',
    tenant_phone: '',
  })
  const [status, setStatus] = useState<'idle' | 'loading' | 'success' | 'error'>('idle')
  const [error, setError] = useState('')

  const set = (key: string, value: string) => setForm(f => ({ ...f, [key]: value }))

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    setStatus('loading')
    setError('')
    try {
      const res = await fetch(`/api/submit/${token}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(form),
      })
      const data = await res.json()
      if (!res.ok) throw new Error(data.error || 'Failed to submit request')
      setStatus('success')
    } catch (err) {
      setStatus('error')
      setError(err instanceof Error ? err.message : 'An error occurred')
    }
  }

  if (status === 'success') {
    return (
      <div className="bg-white rounded-xl border border-gray-200 p-8 text-center">
        <div className="w-12 h-12 rounded-full bg-green-50 flex items-center justify-center mx-auto mb-4">
          <CheckCircle2 className="w-6 h-6 text-green-600" />
        </div>
        <h2 className="text-lg font-semibold text-gray-900">Request submitted</h2>
        <p className="text-sm text-gray-500 mt-2">
          Thanks! Your maintenance request has been received and your property manager has been notified.
          {form.tenant_email && ' You’ll get email updates as it’s handled.'}
        </p>
        <button
          onClick={() => {
            setForm({
              title: '', description: '', urgency: 'medium', property_hint: '',
              unit_number: '', tenant_name: '', tenant_email: '', tenant_phone: '',
            })
            setStatus('idle')
          }}
          className="mt-5 text-sm font-medium text-indigo-600 hover:text-indigo-700"
        >
          Submit another request
        </button>
      </div>
    )
  }

  const loading = status === 'loading'

  return (
    <form onSubmit={handleSubmit} className="space-y-5">
      <div className="bg-white rounded-xl border border-gray-200 p-6 space-y-4">
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1.5">What&apos;s the issue? *</label>
          <Input
            value={form.title}
            onChange={e => set('title', e.target.value)}
            placeholder="e.g. Water leaking under kitchen sink"
            required
            disabled={loading}
          />
        </div>

        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1.5">Describe it *</label>
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
          <label className="block text-sm font-medium text-gray-700 mb-1.5">How urgent is it?</label>
          <div className="space-y-2">
            {URGENCY_OPTIONS.map(opt => (
              <label key={opt.value} className="flex items-start gap-3 cursor-pointer">
                <input
                  type="radio"
                  name="urgency"
                  value={opt.value}
                  checked={form.urgency === opt.value}
                  onChange={e => set('urgency', e.target.value)}
                  className="mt-0.5"
                  disabled={loading}
                />
                <span className="text-sm text-gray-700">{opt.label}</span>
              </label>
            ))}
          </div>
        </div>
      </div>

      <div className="bg-white rounded-xl border border-gray-200 p-6 space-y-4">
        <h2 className="text-sm font-semibold text-gray-700">Where are you?</h2>
        <div className="grid grid-cols-2 gap-4">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1.5">Property / address</label>
            <Input
              value={form.property_hint}
              onChange={e => set('property_hint', e.target.value)}
              placeholder="e.g. 120 Maple St"
              disabled={loading}
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1.5">Unit</label>
            <Input
              value={form.unit_number}
              onChange={e => set('unit_number', e.target.value)}
              placeholder="e.g. 4B"
              disabled={loading}
            />
          </div>
        </div>
      </div>

      <div className="bg-white rounded-xl border border-gray-200 p-6 space-y-4">
        <h2 className="text-sm font-semibold text-gray-700">Your contact info</h2>
        <p className="text-xs text-gray-500 -mt-2">So your property manager can reach you and send updates.</p>
        <div className="grid grid-cols-2 gap-4">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1.5">Name</label>
            <Input
              value={form.tenant_name}
              onChange={e => set('tenant_name', e.target.value)}
              placeholder="Jane Smith"
              disabled={loading}
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1.5">Phone</label>
            <Input
              value={form.tenant_phone}
              onChange={e => set('tenant_phone', e.target.value)}
              placeholder="+1 (555) 000-0000"
              disabled={loading}
            />
          </div>
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1.5">Email</label>
          <Input
            type="email"
            value={form.tenant_email}
            onChange={e => set('tenant_email', e.target.value)}
            placeholder="you@email.com"
            disabled={loading}
          />
        </div>
      </div>

      {error && <p className="text-sm text-red-600">{error}</p>}

      <Button
        type="submit"
        disabled={loading || !form.title || !form.description}
        className="w-full bg-indigo-600 hover:bg-indigo-700 text-white py-3"
      >
        {loading ? 'Submitting...' : 'Submit request'}
      </Button>
    </form>
  )
}

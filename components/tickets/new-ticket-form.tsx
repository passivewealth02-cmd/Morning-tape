'use client'

import { useState } from 'react'
import { useRouter } from 'next/navigation'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Textarea } from '@/components/ui/textarea'
import type { Property, Vendor } from '@/lib/db'
import { Zap } from 'lucide-react'

interface NewTicketFormProps {
  properties: Property[]
  vendors: Vendor[]
}

const URGENCY_OPTIONS = [
  { value: 'low', label: 'Low — Minor issue, schedule within a week' },
  { value: 'medium', label: 'Medium — Should be resolved within 48-72h' },
  { value: 'high', label: 'High — Significant impact, same-day response' },
  { value: 'emergency', label: 'Emergency — Immediate safety risk' },
]

export function NewTicketForm({ properties }: NewTicketFormProps) {
  const router = useRouter()
  const [form, setForm] = useState({
    title: '',
    description: '',
    urgency: 'medium',
    property_id: '',
    unit_number: '',
    tenant_name: '',
    tenant_email: '',
    tenant_phone: '',
  })
  const [status, setStatus] = useState<'idle' | 'loading' | 'error'>('idle')
  const [error, setError] = useState('')

  const set = (key: string, value: string) => setForm(f => ({ ...f, [key]: value }))

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    setStatus('loading')
    setError('')

    try {
      const res = await fetch('/api/tickets', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          ...form,
          property_id: form.property_id || null,
          unit_number: form.unit_number || null,
          tenant_name: form.tenant_name || null,
          tenant_email: form.tenant_email || null,
          tenant_phone: form.tenant_phone || null,
        }),
      })

      if (!res.ok) {
        const data = await res.json()
        throw new Error(data.error || 'Failed to create ticket')
      }

      const ticket = await res.json()
      router.push(`/dashboard/tickets/${ticket.id}`)
    } catch (err) {
      setStatus('error')
      setError(err instanceof Error ? err.message : 'An error occurred')
    }
  }

  return (
    <form onSubmit={handleSubmit} className="space-y-6">
      {/* Issue details */}
      <div className="bg-white rounded-xl border border-gray-200 p-6 space-y-4">
        <h2 className="text-sm font-semibold text-gray-700">Issue Details</h2>

        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1.5">Title *</label>
          <Input
            value={form.title}
            onChange={e => set('title', e.target.value)}
            placeholder="e.g. Water leaking under kitchen sink"
            required
            disabled={status === 'loading'}
          />
        </div>

        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1.5">Description *</label>
          <Textarea
            value={form.description}
            onChange={e => set('description', e.target.value)}
            placeholder="Describe the issue in detail. Include when it started, how severe it is, and any steps already taken."
            required
            rows={4}
            disabled={status === 'loading'}
          />
        </div>

        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1.5">Urgency</label>
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
                  disabled={status === 'loading'}
                />
                <span className="text-sm text-gray-700">{opt.label}</span>
              </label>
            ))}
          </div>
        </div>
      </div>

      {/* Property */}
      <div className="bg-white rounded-xl border border-gray-200 p-6 space-y-4">
        <h2 className="text-sm font-semibold text-gray-700">Location</h2>

        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1.5">Property</label>
          <select
            value={form.property_id}
            onChange={e => set('property_id', e.target.value)}
            className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-white"
            disabled={status === 'loading'}
          >
            <option value="">Select a property (optional)</option>
            {properties.map(p => (
              <option key={p.id} value={p.id}>{p.name}</option>
            ))}
          </select>
        </div>

        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1.5">Unit number</label>
          <Input
            value={form.unit_number}
            onChange={e => set('unit_number', e.target.value)}
            placeholder="e.g. 4B"
            disabled={status === 'loading'}
          />
        </div>
      </div>

      {/* Tenant */}
      <div className="bg-white rounded-xl border border-gray-200 p-6 space-y-4">
        <h2 className="text-sm font-semibold text-gray-700">Tenant Contact (Optional)</h2>

        <div className="grid grid-cols-2 gap-4">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1.5">Name</label>
            <Input
              value={form.tenant_name}
              onChange={e => set('tenant_name', e.target.value)}
              placeholder="Jane Smith"
              disabled={status === 'loading'}
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1.5">Phone</label>
            <Input
              value={form.tenant_phone}
              onChange={e => set('tenant_phone', e.target.value)}
              placeholder="+1 (555) 000-0000"
              disabled={status === 'loading'}
            />
          </div>
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1.5">Email</label>
          <Input
            type="email"
            value={form.tenant_email}
            onChange={e => set('tenant_email', e.target.value)}
            placeholder="tenant@email.com"
            disabled={status === 'loading'}
          />
        </div>
      </div>

      {error && <p className="text-sm text-red-600">{error}</p>}

      <div className="flex items-center gap-3">
        <Button
          type="submit"
          disabled={status === 'loading' || !form.title || !form.description}
          className="bg-indigo-600 hover:bg-indigo-700 text-white"
        >
          {status === 'loading' ? (
            'Creating ticket...'
          ) : (
            <>
              <Zap className="w-4 h-4 mr-1.5" />
              Create ticket with AI
            </>
          )}
        </Button>
        <button
          type="button"
          onClick={() => router.back()}
          className="text-sm text-gray-500 hover:text-gray-700"
        >
          Cancel
        </button>
      </div>

      {status === 'loading' && (
        <p className="text-xs text-indigo-600 flex items-center gap-1.5">
          <Zap className="w-3 h-3" />
          AI is analyzing and categorizing your ticket...
        </p>
      )}
    </form>
  )
}

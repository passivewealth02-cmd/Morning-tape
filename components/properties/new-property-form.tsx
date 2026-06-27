'use client'

import { useState, useMemo } from 'react'
import { useRouter } from 'next/navigation'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { expandUnitSpec, summarizeUnits } from '@/lib/units'

export function NewPropertyForm() {
  const router = useRouter()
  const [form, setForm] = useState({ name: '', address: '', city: '', province: '', unit_spec: '' })
  const [status, setStatus] = useState<'idle' | 'loading' | 'error'>('idle')
  const [error, setError] = useState('')

  const set = (key: string, value: string) => setForm(f => ({ ...f, [key]: value }))

  const units = useMemo(() => expandUnitSpec(form.unit_spec), [form.unit_spec])

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    setStatus('loading')
    setError('')

    try {
      const res = await fetch('/api/properties', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          name: form.name,
          address: form.address,
          city: form.city || null,
          province: form.province || null,
          unit_spec: form.unit_spec,
        }),
      })

      if (!res.ok) {
        const data = await res.json()
        throw new Error(data.error || 'Failed to add property')
      }

      setForm({ name: '', address: '', city: '', province: '', unit_spec: '' })
      setStatus('idle')
      router.refresh()
    } catch (err) {
      setStatus('error')
      setError(err instanceof Error ? err.message : 'An error occurred')
    }
  }

  return (
    <form onSubmit={handleSubmit} className="space-y-3">
      <div>
        <label className="block text-xs font-medium text-gray-700 mb-1">Property name *</label>
        <Input value={form.name} onChange={e => set('name', e.target.value)} placeholder="Maple Gardens" required disabled={status === 'loading'} className="text-sm" />
      </div>
      <div>
        <label className="block text-xs font-medium text-gray-700 mb-1">Address *</label>
        <Input value={form.address} onChange={e => set('address', e.target.value)} placeholder="123 Main St" required disabled={status === 'loading'} className="text-sm" />
      </div>
      <div className="grid grid-cols-2 gap-2">
        <div>
          <label className="block text-xs font-medium text-gray-700 mb-1">City</label>
          <Input value={form.city} onChange={e => set('city', e.target.value)} placeholder="Toronto" disabled={status === 'loading'} className="text-sm" />
        </div>
        <div>
          <label className="block text-xs font-medium text-gray-700 mb-1">Province/State</label>
          <Input value={form.province} onChange={e => set('province', e.target.value)} placeholder="ON" disabled={status === 'loading'} className="text-sm" />
        </div>
      </div>

      <div>
        <label className="block text-xs font-medium text-gray-700 mb-1">Units <span className="text-gray-400 font-normal">(optional)</span></label>
        <Input
          value={form.unit_spec}
          onChange={e => set('unit_spec', e.target.value)}
          placeholder="e.g. 1-20  or  101-110, 201-210"
          disabled={status === 'loading'}
          className="text-sm"
        />
        <p className="mt-1 text-[11px] text-gray-400 leading-snug">
          Type a range like <span className="font-mono text-gray-500">1-20</span>, multiple ranges
          <span className="font-mono text-gray-500"> 101-110, 201-210</span>, or a list
          <span className="font-mono text-gray-500"> 1A, 1B, PH1</span>. We&apos;ll create them all at once.
        </p>
        {units.length > 0 && (
          <p className="mt-1.5 text-[11px] font-medium text-indigo-600">{summarizeUnits(units)}</p>
        )}
      </div>

      {error && <p className="text-xs text-red-600">{error}</p>}

      <Button type="submit" disabled={status === 'loading' || !form.name || !form.address} className="w-full bg-indigo-600 hover:bg-indigo-700 text-white text-sm">
        {status === 'loading'
          ? 'Adding...'
          : units.length > 0
            ? `Add property + ${units.length} unit${units.length === 1 ? '' : 's'}`
            : 'Add property'}
      </Button>
    </form>
  )
}

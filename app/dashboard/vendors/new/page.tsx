'use client'

import { useState } from 'react'
import { useRouter } from 'next/navigation'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Textarea } from '@/components/ui/textarea'
import Link from 'next/link'
import { ChevronLeft } from 'lucide-react'

const TRADE_TYPES = [
  'plumber', 'electrician', 'hvac_technician', 'handyman', 'painter',
  'roofer', 'locksmith', 'pest_control', 'landscaper', 'cleaner',
  'appliance_repair', 'structural', 'general_contractor', 'other',
]

export default function NewVendorPage() {
  const router = useRouter()
  const [form, setForm] = useState({
    name: '',
    trade_type: '',
    email: '',
    phone: '',
    notes: '',
    insurance_status: 'unknown',
  })
  const [status, setStatus] = useState<'idle' | 'loading' | 'error'>('idle')
  const [error, setError] = useState('')

  const set = (key: string, value: string) => setForm(f => ({ ...f, [key]: value }))

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    setStatus('loading')
    setError('')

    try {
      const res = await fetch('/api/vendors', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          ...form,
          email: form.email || null,
          phone: form.phone || null,
          notes: form.notes || null,
        }),
      })

      if (!res.ok) {
        const data = await res.json()
        throw new Error(data.error || 'Failed to create vendor')
      }

      router.push('/dashboard/vendors')
    } catch (err) {
      setStatus('error')
      setError(err instanceof Error ? err.message : 'An error occurred')
    }
  }

  return (
    <div className="p-6 max-w-xl">
      <Link href="/dashboard/vendors" className="inline-flex items-center gap-1 text-sm text-gray-500 hover:text-gray-700 mb-6">
        <ChevronLeft className="w-4 h-4" />
        Back to vendors
      </Link>

      <h1 className="text-xl font-semibold text-gray-900 mb-6">Add Vendor</h1>

      <form onSubmit={handleSubmit} className="space-y-6">
        <div className="bg-white rounded-xl border border-gray-200 p-6 space-y-4">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1.5">Vendor name *</label>
            <Input value={form.name} onChange={e => set('name', e.target.value)} placeholder="ABC Plumbing Co." required disabled={status === 'loading'} />
          </div>

          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1.5">Trade type *</label>
            <select
              value={form.trade_type}
              onChange={e => set('trade_type', e.target.value)}
              required
              className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-white"
              disabled={status === 'loading'}
            >
              <option value="">Select trade type</option>
              {TRADE_TYPES.map(t => (
                <option key={t} value={t}>{t.replace('_', ' ').replace(/\b\w/g, c => c.toUpperCase())}</option>
              ))}
            </select>
          </div>

          <div className="grid grid-cols-2 gap-4">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1.5">Phone</label>
              <Input value={form.phone} onChange={e => set('phone', e.target.value)} placeholder="+1 (555) 000-0000" disabled={status === 'loading'} />
            </div>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1.5">Email</label>
              <Input type="email" value={form.email} onChange={e => set('email', e.target.value)} placeholder="vendor@email.com" disabled={status === 'loading'} />
            </div>
          </div>

          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1.5">Insurance status</label>
            <select
              value={form.insurance_status}
              onChange={e => set('insurance_status', e.target.value)}
              className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-white"
              disabled={status === 'loading'}
            >
              <option value="unknown">Unknown</option>
              <option value="verified">Verified</option>
              <option value="expired">Expired</option>
            </select>
          </div>

          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1.5">Notes</label>
            <Textarea value={form.notes} onChange={e => set('notes', e.target.value)} placeholder="Any notes about this vendor..." rows={3} disabled={status === 'loading'} />
          </div>
        </div>

        {error && <p className="text-sm text-red-600">{error}</p>}

        <div className="flex items-center gap-3">
          <Button type="submit" disabled={status === 'loading' || !form.name || !form.trade_type} className="bg-indigo-600 hover:bg-indigo-700 text-white">
            {status === 'loading' ? 'Adding vendor...' : 'Add vendor'}
          </Button>
          <button type="button" onClick={() => router.back()} className="text-sm text-gray-500 hover:text-gray-700">
            Cancel
          </button>
        </div>
      </form>
    </div>
  )
}

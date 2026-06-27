'use client'

import { useState, useMemo } from 'react'
import { useRouter } from 'next/navigation'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Trash2, DoorOpen } from 'lucide-react'
import { expandUnitSpec, summarizeUnits } from '@/lib/units'
import type { Unit } from '@/lib/db'

export function UnitManager({ propertyId, units }: { propertyId: string; units: Unit[] }) {
  const router = useRouter()
  const [spec, setSpec] = useState('')
  const [status, setStatus] = useState<'idle' | 'loading' | 'error'>('idle')
  const [error, setError] = useState('')
  const [deletingId, setDeletingId] = useState<string | null>(null)

  const parsed = useMemo(() => expandUnitSpec(spec), [spec])

  const addUnits = async (e: React.FormEvent) => {
    e.preventDefault()
    if (parsed.length === 0) return
    setStatus('loading')
    setError('')
    try {
      const res = await fetch(`/api/properties/${propertyId}/units`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ unit_spec: spec }),
      })
      if (!res.ok) {
        const data = await res.json()
        throw new Error(data.error || 'Failed to add units')
      }
      setSpec('')
      setStatus('idle')
      router.refresh()
    } catch (err) {
      setStatus('error')
      setError(err instanceof Error ? err.message : 'An error occurred')
    }
  }

  const deleteUnit = async (unitId: string) => {
    setDeletingId(unitId)
    try {
      const res = await fetch(`/api/properties/${propertyId}/units/${unitId}`, { method: 'DELETE' })
      if (!res.ok) throw new Error('Failed to delete unit')
      router.refresh()
    } catch {
      // surface nothing intrusive; leave the row in place on failure
    } finally {
      setDeletingId(null)
    }
  }

  return (
    <div className="grid lg:grid-cols-3 gap-6">
      {/* Unit list */}
      <div className="lg:col-span-2">
        <div className="flex items-center justify-between mb-3">
          <h2 className="text-sm font-semibold text-gray-700">Units ({units.length})</h2>
        </div>
        {units.length === 0 ? (
          <div className="bg-white rounded-xl border border-dashed border-gray-200 p-10 text-center">
            <DoorOpen className="w-7 h-7 text-gray-300 mx-auto mb-2" />
            <p className="text-gray-500 text-sm">No units yet</p>
            <p className="text-xs text-gray-400 mt-0.5">Add a range like 1-20 on the right to create them in one go.</p>
          </div>
        ) : (
          <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
            {units.map(u => (
              <div key={u.id} className="flex items-center justify-between bg-white rounded-lg border border-gray-200 px-3 py-2">
                <span className="text-sm text-gray-800 truncate">{u.unit_number}</span>
                <button
                  type="button"
                  onClick={() => deleteUnit(u.id)}
                  disabled={deletingId === u.id}
                  className="text-gray-300 hover:text-red-500 transition-colors shrink-0"
                  aria-label={`Delete unit ${u.unit_number}`}
                >
                  <Trash2 className="w-3.5 h-3.5" />
                </button>
              </div>
            ))}
          </div>
        )}
      </div>

      {/* Add units */}
      <div>
        <div className="bg-white rounded-xl border border-gray-200 p-5">
          <h2 className="text-sm font-semibold text-gray-700 mb-4">Add units</h2>
          <form onSubmit={addUnits} className="space-y-3">
            <Input
              value={spec}
              onChange={e => setSpec(e.target.value)}
              placeholder="e.g. 1-20  or  101-110, 201-210"
              disabled={status === 'loading'}
              className="text-sm"
            />
            <p className="text-[11px] text-gray-400 leading-snug">
              Ranges <span className="font-mono text-gray-500">1-20</span>, multiple ranges
              <span className="font-mono text-gray-500"> 101-110, 201-210</span>, or a list
              <span className="font-mono text-gray-500"> 1A, 1B, PH1</span>. Duplicates are skipped.
            </p>
            {parsed.length > 0 && (
              <p className="text-[11px] font-medium text-indigo-600">{summarizeUnits(parsed)}</p>
            )}
            {error && <p className="text-xs text-red-600">{error}</p>}
            <Button
              type="submit"
              disabled={status === 'loading' || parsed.length === 0}
              className="w-full bg-indigo-600 hover:bg-indigo-700 text-white text-sm"
            >
              {status === 'loading' ? 'Adding...' : parsed.length > 0 ? `Add ${parsed.length} unit${parsed.length === 1 ? '' : 's'}` : 'Add units'}
            </Button>
          </form>
        </div>
      </div>
    </div>
  )
}

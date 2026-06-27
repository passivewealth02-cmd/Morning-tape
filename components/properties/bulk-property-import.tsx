'use client'

import { useState, useMemo } from 'react'
import { useRouter } from 'next/navigation'
import { Button } from '@/components/ui/button'
import { expandUnitSpec } from '@/lib/units'

type ParsedRow = {
  name: string
  address: string
  city: string | null
  province: string | null
  units: string[]
  error?: string
}

// Each line: Name | Address | City | Province | Units
// Only Name and Address are required. Units accepts the same syntax as the
// single form (e.g. "1-20" or "101-110, 201-210"); use ; to separate ranges
// inside the units column so it doesn't clash with the | column separator.
function parseLine(line: string): ParsedRow | null {
  const trimmed = line.trim()
  if (!trimmed) return null
  const parts = trimmed.split('|').map(p => p.trim())
  const [name = '', address = '', city = '', province = '', unitSpec = ''] = parts
  const units = expandUnitSpec(unitSpec.replace(/;/g, ','))
  const row: ParsedRow = {
    name,
    address,
    city: city || null,
    province: province || null,
    units,
  }
  if (!name || !address) row.error = 'Needs name and address'
  return row
}

export function BulkPropertyImport() {
  const router = useRouter()
  const [text, setText] = useState('')
  const [status, setStatus] = useState<'idle' | 'loading' | 'error'>('idle')
  const [error, setError] = useState('')

  const rows = useMemo(
    () => text.split('\n').map(parseLine).filter((r): r is ParsedRow => r !== null),
    [text]
  )
  const validRows = rows.filter(r => !r.error)
  const errorRows = rows.filter(r => r.error)
  const totalUnits = validRows.reduce((sum, r) => sum + r.units.length, 0)

  const handleSubmit = async () => {
    if (validRows.length === 0) return
    setStatus('loading')
    setError('')
    try {
      const res = await fetch('/api/properties', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          properties: validRows.map(r => ({
            name: r.name,
            address: r.address,
            city: r.city,
            province: r.province,
            units: r.units,
          })),
        }),
      })
      if (!res.ok) {
        const data = await res.json()
        throw new Error(data.error || 'Failed to import properties')
      }
      setText('')
      setStatus('idle')
      router.refresh()
    } catch (err) {
      setStatus('error')
      setError(err instanceof Error ? err.message : 'An error occurred')
    }
  }

  return (
    <div className="space-y-3">
      <div className="rounded-lg bg-gray-50 border border-gray-100 p-3 text-[11px] text-gray-500 leading-relaxed">
        <p className="font-medium text-gray-700 mb-1">One property per line:</p>
        <p className="font-mono text-gray-600">Name | Address | City | Province | Units</p>
        <p className="mt-1.5">Only name &amp; address are required. In the Units column use <span className="font-mono">;</span> between ranges.</p>
        <p className="mt-1 font-mono text-gray-600">Maple Gardens | 123 Main St | Toronto | ON | 1-20</p>
        <p className="font-mono text-gray-600">Oak Towers | 50 King St | Toronto | ON | 101-110; 201-210</p>
      </div>

      <textarea
        value={text}
        onChange={e => setText(e.target.value)}
        disabled={status === 'loading'}
        rows={8}
        placeholder={'Maple Gardens | 123 Main St | Toronto | ON | 1-20\nOak Towers | 50 King St | Toronto | ON | 101-110; 201-210'}
        className="w-full rounded-md border border-gray-200 p-3 text-sm font-mono focus:outline-none focus:ring-2 focus:ring-indigo-200 focus:border-indigo-300"
      />

      {rows.length > 0 && (
        <div className="text-[11px] space-y-1">
          <p className="font-medium text-indigo-600">
            {validRows.length} propert{validRows.length === 1 ? 'y' : 'ies'} · {totalUnits} unit{totalUnits === 1 ? '' : 's'} ready to import
          </p>
          {errorRows.length > 0 && (
            <p className="text-amber-600">{errorRows.length} line{errorRows.length === 1 ? '' : 's'} skipped (missing name or address)</p>
          )}
        </div>
      )}

      {error && <p className="text-xs text-red-600">{error}</p>}

      <Button
        type="button"
        onClick={handleSubmit}
        disabled={status === 'loading' || validRows.length === 0}
        className="w-full bg-indigo-600 hover:bg-indigo-700 text-white text-sm"
      >
        {status === 'loading'
          ? 'Importing...'
          : validRows.length > 0
            ? `Import ${validRows.length} propert${validRows.length === 1 ? 'y' : 'ies'}`
            : 'Import properties'}
      </Button>
    </div>
  )
}

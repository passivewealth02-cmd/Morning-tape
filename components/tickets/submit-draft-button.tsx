'use client'

import { useState } from 'react'
import { useRouter } from 'next/navigation'
import { Zap } from 'lucide-react'

export function SubmitDraftButton({ ticketId }: { ticketId: string }) {
  const router = useRouter()
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState('')

  const submit = async () => {
    setBusy(true)
    setError('')
    try {
      const res = await fetch(`/api/tickets/${ticketId}/submit`, { method: 'POST' })
      if (!res.ok) {
        const data = await res.json().catch(() => ({}))
        throw new Error(data.error || 'Failed to submit draft')
      }
      router.refresh()
    } catch (err) {
      setBusy(false)
      setError(err instanceof Error ? err.message : 'Something went wrong')
    }
  }

  return (
    <div className="flex flex-col items-end gap-1">
      <button
        type="button"
        onClick={submit}
        disabled={busy}
        className="inline-flex items-center gap-1.5 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-medium px-4 py-2 transition-colors disabled:opacity-60"
      >
        <Zap className="w-4 h-4" />
        {busy ? 'Submitting…' : 'Submit ticket'}
      </button>
      {error && <span className="text-xs text-red-600">{error}</span>}
    </div>
  )
}

'use client'

import { useState } from 'react'
import { useRouter } from 'next/navigation'
import { Trash2 } from 'lucide-react'

export function DeleteButton({
  endpoint,
  redirectTo,
  label = 'Delete',
  confirmLabel = 'Confirm delete',
  size = 'md',
}: {
  endpoint: string
  redirectTo?: string
  label?: string
  confirmLabel?: string
  size?: 'sm' | 'md'
}) {
  const router = useRouter()
  const [confirming, setConfirming] = useState(false)
  const [busy, setBusy] = useState(false)
  const [failed, setFailed] = useState(false)

  const pad = size === 'sm' ? 'px-2.5 py-1 text-xs' : 'px-3 py-2 text-sm'

  const run = async () => {
    setBusy(true)
    setFailed(false)
    try {
      const res = await fetch(endpoint, { method: 'DELETE' })
      if (!res.ok) throw new Error('failed')
      if (redirectTo) router.push(redirectTo)
      router.refresh()
    } catch {
      setBusy(false)
      setConfirming(false)
      setFailed(true)
    }
  }

  if (confirming) {
    return (
      <span className="inline-flex items-center gap-1.5">
        <button
          type="button"
          onClick={run}
          disabled={busy}
          className={`inline-flex items-center gap-1.5 rounded-lg bg-red-600 hover:bg-red-700 text-white font-medium transition-colors ${pad}`}
        >
          <Trash2 className="w-3.5 h-3.5" />
          {busy ? 'Deleting…' : confirmLabel}
        </button>
        <button
          type="button"
          onClick={() => setConfirming(false)}
          disabled={busy}
          className={`rounded-lg border border-gray-200 hover:bg-gray-50 text-gray-600 font-medium transition-colors ${pad}`}
        >
          Cancel
        </button>
      </span>
    )
  }

  return (
    <button
      type="button"
      onClick={() => setConfirming(true)}
      className={`inline-flex items-center gap-1.5 rounded-lg border border-gray-200 text-gray-500 hover:text-red-600 hover:border-red-200 hover:bg-red-50 font-medium transition-colors ${pad}`}
    >
      <Trash2 className="w-3.5 h-3.5" />
      {failed ? 'Try again' : label}
    </button>
  )
}

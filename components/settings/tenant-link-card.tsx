'use client'

import { useState } from 'react'
import { Share2, Copy, Check, ExternalLink } from 'lucide-react'

export function TenantLinkCard({ url }: { url: string }) {
  const [copied, setCopied] = useState(false)

  const copy = async () => {
    try {
      await navigator.clipboard.writeText(url)
      setCopied(true)
      setTimeout(() => setCopied(false), 1500)
    } catch {
      // ignore
    }
  }

  return (
    <section className="bg-white rounded-lg border border-gray-200 p-6 mb-4">
      <div className="flex items-start gap-3 mb-4">
        <div className="w-8 h-8 rounded-md bg-indigo-50 flex items-center justify-center shrink-0">
          <Share2 className="w-4 h-4 text-indigo-600" />
        </div>
        <div>
          <h2 className="text-sm font-semibold text-gray-900">Tenant request link</h2>
          <p className="text-xs text-gray-500 mt-0.5">
            Share this link with tenants so they can submit maintenance requests directly. Each submission becomes an AI-categorized ticket.
          </p>
        </div>
      </div>

      <label className="block text-xs font-medium text-gray-700 mb-1.5">Public form URL</label>
      <div className="flex gap-2">
        <input
          readOnly
          value={url}
          className="flex-1 text-xs font-mono px-3 py-2 bg-gray-50 border border-gray-200 rounded-md text-gray-700"
          onFocus={e => e.currentTarget.select()}
        />
        <button
          onClick={copy}
          className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-medium text-gray-700 bg-white border border-gray-200 rounded-md hover:bg-gray-50 transition-colors"
        >
          {copied ? <Check className="w-3.5 h-3.5 text-green-600" /> : <Copy className="w-3.5 h-3.5" />}
          {copied ? 'Copied' : 'Copy'}
        </button>
        <a
          href={url}
          target="_blank"
          rel="noopener noreferrer"
          className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-medium text-gray-700 bg-white border border-gray-200 rounded-md hover:bg-gray-50 transition-colors"
        >
          <ExternalLink className="w-3.5 h-3.5" />
          Open
        </a>
      </div>

      <p className="mt-3 text-xs text-gray-500">
        Tip: add this link to your tenant portal, lease docs, or a QR code posted in your buildings.
      </p>
    </section>
  )
}

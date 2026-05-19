'use client'

import { useState } from 'react'
import { Mail, Copy, Check } from 'lucide-react'

export function InboxWebhookCard({ url }: { url: string }) {
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
    <section className="bg-white rounded-lg border border-gray-200 p-6">
      <div className="flex items-start gap-3 mb-4">
        <div className="w-8 h-8 rounded-md bg-indigo-50 flex items-center justify-center shrink-0">
          <Mail className="w-4 h-4 text-indigo-600" />
        </div>
        <div>
          <h2 className="text-sm font-semibold text-gray-900">Email intake</h2>
          <p className="text-xs text-gray-500 mt-0.5">
            Forward maintenance emails to this webhook. AI parses the message and creates a ticket automatically.
          </p>
        </div>
      </div>

      <label className="block text-xs font-medium text-gray-700 mb-1.5">
        Webhook URL
      </label>
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
      </div>

      <div className="mt-4 p-3 bg-gray-50 rounded-md text-xs text-gray-600 space-y-2">
        <p className="font-medium text-gray-700">How to use:</p>
        <ol className="list-decimal list-inside space-y-1">
          <li>Connect this URL to Zapier (Email Parser → Webhook) or Mailgun/Resend inbound routes</li>
          <li>POST JSON: <code className="bg-white px-1 py-0.5 rounded">{`{ from, subject, text }`}</code></li>
          <li>AI parses the email, extracts the issue, creates and auto-assigns a ticket</li>
        </ol>
      </div>
    </section>
  )
}

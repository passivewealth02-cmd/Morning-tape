'use client'

import { useState } from 'react'
import { Download, Printer, Mail, Link2, Check } from 'lucide-react'

export function QrActions({
  imageUrl,
  downloadName,
  submitUrl,
  title,
}: {
  imageUrl: string
  downloadName: string
  submitUrl: string
  title: string
}) {
  const [copied, setCopied] = useState(false)
  const [busy, setBusy] = useState(false)

  const download = async () => {
    setBusy(true)
    try {
      const res = await fetch(imageUrl)
      const blob = await res.blob()
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = downloadName
      document.body.appendChild(a)
      a.click()
      a.remove()
      URL.revokeObjectURL(url)
    } catch {
      // Fallback: let the QR service serve it as an attachment in a new tab.
      window.open(`${imageUrl}&download=1`, '_blank')
    } finally {
      setBusy(false)
    }
  }

  const email = () => {
    const subject = `Maintenance request QR — ${title}`
    const body =
      `Scan this QR code or open the link below to submit a maintenance request for ${title}:\n\n` +
      `${submitUrl}\n\n` +
      `Tip: download the QR image from your Maintena dashboard to print or post it on the unit door.`
    window.location.href = `mailto:?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`
  }

  const copy = async () => {
    try {
      await navigator.clipboard.writeText(submitUrl)
      setCopied(true)
      setTimeout(() => setCopied(false), 1500)
    } catch {
      // clipboard unavailable — no-op
    }
  }

  const btn =
    'inline-flex items-center gap-1.5 text-sm font-medium px-3.5 py-2 rounded-lg transition-colors'

  return (
    <div className="flex flex-wrap justify-center gap-2 print:hidden">
      <button onClick={download} disabled={busy} className={`${btn} bg-indigo-600 hover:bg-indigo-700 text-white`}>
        <Download className="w-4 h-4" />
        {busy ? 'Preparing…' : 'Download PNG'}
      </button>
      <button onClick={() => window.print()} className={`${btn} border border-gray-200 hover:bg-gray-50 text-gray-700`}>
        <Printer className="w-4 h-4" />
        Print
      </button>
      <button onClick={email} className={`${btn} border border-gray-200 hover:bg-gray-50 text-gray-700`}>
        <Mail className="w-4 h-4" />
        Email
      </button>
      <button onClick={copy} className={`${btn} border border-gray-200 hover:bg-gray-50 text-gray-700`}>
        {copied ? <Check className="w-4 h-4 text-green-600" /> : <Link2 className="w-4 h-4" />}
        {copied ? 'Copied!' : 'Copy link'}
      </button>
    </div>
  )
}

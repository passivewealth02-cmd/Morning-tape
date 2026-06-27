'use client'

import { Printer } from 'lucide-react'

export function PrintButton() {
  return (
    <button
      type="button"
      onClick={() => window.print()}
      className="inline-flex items-center gap-1.5 text-sm font-medium bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2 rounded-lg transition-colors print:hidden"
    >
      <Printer className="w-4 h-4" />
      Print / Save as PDF
    </button>
  )
}

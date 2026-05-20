'use client'

import Link from 'next/link'
import { useState } from 'react'
import { ArrowRight } from 'lucide-react'
import { HowItWorksModal } from './how-it-works-modal'

export function HeroCTA() {
  const [modalOpen, setModalOpen] = useState(false)

  return (
    <>
      <div className="flex flex-col sm:flex-row gap-3">
        <Link
          href="/login"
          className="inline-flex items-center gap-2 bg-gray-900 text-white px-6 py-3 rounded-lg font-medium hover:bg-gray-700 transition-colors"
        >
          Start for free
          <ArrowRight className="w-4 h-4" />
        </Link>
        <button
          onClick={() => setModalOpen(true)}
          className="inline-flex items-center gap-2 bg-gray-50 text-gray-700 px-6 py-3 rounded-lg font-medium hover:bg-gray-100 transition-colors border border-gray-200"
        >
          See how it works
        </button>
      </div>
      <HowItWorksModal open={modalOpen} onClose={() => setModalOpen(false)} />
    </>
  )
}

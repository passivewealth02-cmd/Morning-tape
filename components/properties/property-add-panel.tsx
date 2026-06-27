'use client'

import { useState } from 'react'
import { Plus, Upload } from 'lucide-react'
import { NewPropertyForm } from './new-property-form'
import { BulkPropertyImport } from './bulk-property-import'

export function PropertyAddPanel() {
  const [mode, setMode] = useState<'single' | 'bulk'>('single')

  return (
    <div className="bg-white rounded-xl border border-gray-200 p-5">
      <div className="flex items-center gap-1 mb-4 bg-gray-100 rounded-lg p-1">
        <button
          type="button"
          onClick={() => setMode('single')}
          className={`flex-1 flex items-center justify-center gap-1.5 text-xs font-medium rounded-md py-1.5 transition-colors ${
            mode === 'single' ? 'bg-white text-gray-900 shadow-sm' : 'text-gray-500 hover:text-gray-700'
          }`}
        >
          <Plus className="w-3.5 h-3.5" />
          Add one
        </button>
        <button
          type="button"
          onClick={() => setMode('bulk')}
          className={`flex-1 flex items-center justify-center gap-1.5 text-xs font-medium rounded-md py-1.5 transition-colors ${
            mode === 'bulk' ? 'bg-white text-gray-900 shadow-sm' : 'text-gray-500 hover:text-gray-700'
          }`}
        >
          <Upload className="w-3.5 h-3.5" />
          Bulk import
        </button>
      </div>

      {mode === 'single' ? <NewPropertyForm /> : <BulkPropertyImport />}
    </div>
  )
}

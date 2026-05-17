'use client'

import { useState } from 'react'
import { useRouter } from 'next/navigation'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'

export default function OnboardingPage() {
  const router = useRouter()
  const [orgName, setOrgName] = useState('')
  const [status, setStatus] = useState<'idle' | 'loading' | 'error'>('idle')
  const [errorMessage, setErrorMessage] = useState('')

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    setStatus('loading')
    setErrorMessage('')

    try {
      const response = await fetch('/api/onboarding', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ organization_name: orgName }),
      })

      const data = await response.json()

      if (!response.ok) {
        throw new Error(data.error || 'Failed to create workspace')
      }

      router.push('/dashboard')
    } catch (error) {
      setStatus('error')
      setErrorMessage(error instanceof Error ? error.message : 'An error occurred')
    }
  }

  return (
    <div className="min-h-screen flex items-center justify-center px-6 bg-gray-50">
      <div className="max-w-sm w-full">
        <div className="text-center mb-8">
          <h1 className="text-2xl font-semibold text-gray-900">Set up your workspace</h1>
          <p className="text-gray-500 text-sm mt-2">Tell us the name of your organization to get started.</p>
        </div>

        <div className="bg-white rounded-xl border border-gray-200 p-8 shadow-sm">
          <form onSubmit={handleSubmit} className="space-y-4">
            <div>
              <label htmlFor="org-name" className="block text-sm font-medium text-gray-700 mb-1.5">
                Organization name
              </label>
              <Input
                id="org-name"
                type="text"
                value={orgName}
                onChange={e => setOrgName(e.target.value)}
                placeholder="Acme Property Management"
                required
                disabled={status === 'loading'}
                className="w-full"
              />
            </div>

            {status === 'error' && (
              <p className="text-sm text-red-600">{errorMessage}</p>
            )}

            <Button
              type="submit"
              disabled={status === 'loading' || !orgName.trim()}
              className="w-full bg-gray-900 hover:bg-gray-700 text-white"
            >
              {status === 'loading' ? 'Creating workspace...' : 'Continue'}
            </Button>
          </form>
        </div>
      </div>
    </div>
  )
}

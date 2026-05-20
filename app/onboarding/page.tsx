'use client'

import { useState } from 'react'
import { Check } from 'lucide-react'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'

type Plan = 'starter' | 'growth' | 'pro'

const PLANS: { id: Plan; name: string; price: string; description: string; features: string[]; highlighted?: boolean }[] = [
  {
    id: 'starter',
    name: 'Starter',
    price: '$99/mo',
    description: 'Small portfolios',
    features: ['100 tickets/month', '1 property', '5 vendors', 'Email notifications'],
  },
  {
    id: 'growth',
    name: 'Growth',
    price: '$299/mo',
    description: 'Growing teams',
    features: ['Unlimited tickets', '10 properties', 'Unlimited vendors', 'AI dispatch', 'SLA tracking'],
    highlighted: true,
  },
  {
    id: 'pro',
    name: 'Pro',
    price: '$599/mo',
    description: 'Large portfolios',
    features: ['Unlimited everything', 'White-label', 'Dedicated CSM', 'API access'],
  },
]

export default function OnboardingPage() {
  const [step, setStep] = useState<'name' | 'plan'>('name')
  const [orgName, setOrgName] = useState('')
  const [selectedPlan, setSelectedPlan] = useState<Plan>('growth')
  const [status, setStatus] = useState<'idle' | 'loading' | 'error'>('idle')
  const [errorMessage, setErrorMessage] = useState('')

  const handleNameSubmit = (e: React.FormEvent) => {
    e.preventDefault()
    if (orgName.trim()) setStep('plan')
  }

  const handlePlanSubmit = async () => {
    setStatus('loading')
    setErrorMessage('')
    try {
      // Step 1: create org
      const onboardRes = await fetch('/api/onboarding', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ organization_name: orgName, plan: selectedPlan }),
      })
      const onboardData = await onboardRes.json()
      if (!onboardRes.ok) throw new Error(onboardData.error || 'Failed to create workspace')

      // Step 2: start Stripe checkout with 14-day trial
      const checkoutRes = await fetch('/api/billing/checkout', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ plan: selectedPlan }),
      })
      const checkoutData = await checkoutRes.json()
      if (!checkoutRes.ok) throw new Error(checkoutData.error || 'Failed to start checkout')

      window.location.href = checkoutData.url
    } catch (error) {
      setStatus('error')
      setErrorMessage(error instanceof Error ? error.message : 'An error occurred')
    }
  }

  if (step === 'name') {
    return (
      <div className="min-h-screen flex items-center justify-center px-6 bg-gray-50">
        <div className="max-w-sm w-full">
          <div className="text-center mb-8">
            <h1 className="text-2xl font-semibold text-gray-900">Set up your workspace</h1>
            <p className="text-gray-500 text-sm mt-2">Tell us the name of your organization to get started.</p>
          </div>
          <div className="bg-white rounded-xl border border-gray-200 p-8 shadow-sm">
            <form onSubmit={handleNameSubmit} className="space-y-4">
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
                  className="w-full"
                />
              </div>
              <Button
                type="submit"
                disabled={!orgName.trim()}
                className="w-full bg-gray-900 hover:bg-gray-700 text-white"
              >
                Continue
              </Button>
            </form>
          </div>
        </div>
      </div>
    )
  }

  return (
    <div className="min-h-screen flex items-center justify-center px-6 py-12 bg-gray-50">
      <div className="max-w-2xl w-full">
        <div className="text-center mb-8">
          <h1 className="text-2xl font-semibold text-gray-900">Choose your plan</h1>
          <p className="text-gray-500 text-sm mt-2">
            Free for 14 days. Your card will be charged after the trial ends. Cancel anytime.
          </p>
        </div>

        <div className="grid sm:grid-cols-3 gap-4 mb-6">
          {PLANS.map(plan => {
            const isSelected = selectedPlan === plan.id
            return (
              <button
                key={plan.id}
                onClick={() => setSelectedPlan(plan.id)}
                className={`text-left rounded-xl border p-5 transition-all ${
                  isSelected
                    ? 'border-indigo-400 bg-indigo-50 ring-2 ring-indigo-300'
                    : plan.highlighted
                      ? 'border-indigo-200 bg-white hover:border-indigo-300'
                      : 'border-gray-200 bg-white hover:border-gray-300'
                }`}
              >
                <div className="flex items-start justify-between mb-1">
                  <span className="text-sm font-semibold text-gray-900">{plan.name}</span>
                  {isSelected && (
                    <span className="w-5 h-5 rounded-full bg-indigo-600 flex items-center justify-center shrink-0">
                      <Check className="w-3 h-3 text-white" />
                    </span>
                  )}
                </div>
                <p className="text-xs text-gray-500 mb-2">{plan.description}</p>
                <p className="text-lg font-bold text-gray-900 mb-3">{plan.price}</p>
                <ul className="space-y-1.5">
                  {plan.features.map(f => (
                    <li key={f} className="flex items-start gap-1.5 text-xs text-gray-600">
                      <Check className="w-3 h-3 mt-0.5 shrink-0 text-indigo-500" />
                      {f}
                    </li>
                  ))}
                </ul>
              </button>
            )
          })}
        </div>

        {status === 'error' && (
          <p className="text-sm text-red-600 mb-4 text-center">{errorMessage}</p>
        )}

        <Button
          onClick={handlePlanSubmit}
          disabled={status === 'loading'}
          className="w-full bg-indigo-600 hover:bg-indigo-700 text-white py-3 text-sm font-medium"
        >
          {status === 'loading' ? 'Setting up...' : `Start 14-day free trial — ${PLANS.find(p => p.id === selectedPlan)?.price} after`}
        </Button>

        <p className="text-center text-xs text-gray-400 mt-3">
          You won&apos;t be charged until your trial ends. Cancel anytime before day 14.
        </p>

        <button
          onClick={() => setStep('name')}
          className="mt-4 w-full text-center text-xs text-gray-400 hover:text-gray-600"
        >
          ← Back
        </button>
      </div>
    </div>
  )
}

'use client'

import { useState } from 'react'
import { useRouter } from 'next/navigation'
import { Check, Sparkles } from 'lucide-react'
import type { OrganizationPlan } from '@/lib/db'

type Plan = {
  id: OrganizationPlan
  name: string
  price: string
  period: string
  description: string
  features: string[]
  highlighted?: boolean
}

const PLANS: Plan[] = [
  {
    id: 'starter',
    name: 'Starter',
    price: '$99',
    period: '/mo',
    description: 'Small portfolios',
    features: ['100 tickets/mo', '1 property', '5 vendors', 'Email notifications'],
  },
  {
    id: 'growth',
    name: 'Growth',
    price: '$299',
    period: '/mo',
    description: 'Growing teams',
    features: ['Unlimited tickets', '10 properties', 'Unlimited vendors', 'AI dispatch', 'SLA tracking'],
    highlighted: true,
  },
  {
    id: 'pro',
    name: 'Pro',
    price: '$599',
    period: '/mo',
    description: 'Large portfolios',
    features: ['Unlimited everything', 'White-label', 'Dedicated CSM', 'API access'],
  },
]

export type SubscriptionInfo = {
  status: string
  cancelAtPeriodEnd: boolean
  endsAt: string | null
}

export function PlanCard({
  currentPlan,
  canManage,
  hasSubscription,
  subscription,
}: {
  currentPlan: OrganizationPlan
  canManage: boolean
  hasSubscription: boolean
  subscription?: SubscriptionInfo | null
}) {
  const router = useRouter()
  const [pending, setPending] = useState<OrganizationPlan | 'portal' | 'cancel' | null>(null)
  const [error, setError] = useState('')
  const [confirmingCancel, setConfirmingCancel] = useState(false)

  const isTrialing = subscription?.status === 'trialing'
  const scheduledToCancel = subscription?.cancelAtPeriodEnd ?? false
  const endsAtLabel = subscription?.endsAt
    ? new Date(subscription.endsAt).toLocaleDateString(undefined, { month: 'long', day: 'numeric', year: 'numeric' })
    : null

  const setCancellation = async (reactivate: boolean) => {
    setPending('cancel')
    setError('')
    try {
      const res = await fetch('/api/billing/cancel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ reactivate }),
      })
      const data = await res.json()
      if (!res.ok) throw new Error(data.error || 'Could not update subscription')
      setConfirmingCancel(false)
      router.refresh()
    } catch (err) {
      setError(err instanceof Error ? err.message : 'An error occurred')
    } finally {
      setPending(null)
    }
  }

  const handleSelect = async (plan: OrganizationPlan) => {
    if (plan === currentPlan || pending) return
    setPending(plan)
    setError('')
    try {
      const res = await fetch('/api/billing/checkout', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ plan }),
      })
      const data = await res.json()
      if (!res.ok) throw new Error(data.error || 'Failed to start checkout')
      if (data.url) {
        window.location.href = data.url
        return
      }
      router.refresh()
    } catch (err) {
      setError(err instanceof Error ? err.message : 'An error occurred')
      setPending(null)
    }
  }

  const handleManage = async () => {
    setPending('portal')
    setError('')
    try {
      const res = await fetch('/api/billing/portal', { method: 'POST' })
      const data = await res.json()
      if (!res.ok) throw new Error(data.error || 'Could not open billing portal')
      if (data.url) window.location.href = data.url
    } catch (err) {
      setError(err instanceof Error ? err.message : 'An error occurred')
      setPending(null)
    }
  }

  return (
    <section className="bg-white rounded-lg border border-gray-200 p-6">
      <div className="flex items-start gap-3 mb-4">
        <div className="w-8 h-8 rounded-md bg-indigo-50 flex items-center justify-center shrink-0">
          <Sparkles className="w-4 h-4 text-indigo-600" />
        </div>
        <div className="flex-1">
          <div className="flex items-center justify-between">
            <h2 className="text-sm font-semibold text-gray-900">Plan</h2>
            <span className="inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium bg-indigo-50 text-indigo-700 capitalize">
              {currentPlan}
            </span>
          </div>
          <p className="text-xs text-gray-500 mt-0.5">
            {currentPlan === 'trial'
              ? 'You are on a free trial. Pick a plan to unlock more.'
              : 'Change or upgrade your plan at any time.'}
          </p>
        </div>
      </div>

      {error && (
        <p className="text-sm text-red-600 mb-3">{error}</p>
      )}

      <div className="grid sm:grid-cols-3 gap-3">
        {PLANS.map(plan => {
          const isCurrent = currentPlan === plan.id
          const isPending = pending === plan.id
          return (
            <div
              key={plan.id}
              className={`rounded-lg border p-4 flex flex-col ${
                isCurrent
                  ? 'border-indigo-300 bg-indigo-50 ring-1 ring-indigo-200'
                  : plan.highlighted
                    ? 'border-indigo-200 bg-white'
                    : 'border-gray-200 bg-white'
              }`}
            >
              <div className="flex items-baseline justify-between mb-1">
                <h3 className="text-sm font-semibold text-gray-900">{plan.name}</h3>
                {plan.highlighted && !isCurrent && (
                  <span className="text-[10px] uppercase tracking-wide font-semibold text-indigo-600">Popular</span>
                )}
              </div>
              <p className="text-xs text-gray-500 mb-3">{plan.description}</p>
              <div className="flex items-baseline gap-1 mb-3">
                <span className="text-xl font-semibold text-gray-900">{plan.price}</span>
                <span className="text-xs text-gray-500">{plan.period}</span>
              </div>
              <ul className="space-y-1.5 mb-4 flex-1">
                {plan.features.map(f => (
                  <li key={f} className="flex items-start gap-1.5 text-xs text-gray-600">
                    <Check className="w-3 h-3 mt-0.5 shrink-0 text-indigo-600" />
                    {f}
                  </li>
                ))}
              </ul>
              <button
                disabled={!canManage || isCurrent || pending !== null}
                onClick={() => handleSelect(plan.id)}
                className={`w-full text-xs font-medium px-3 py-2 rounded-md transition-colors ${
                  isCurrent
                    ? 'bg-indigo-100 text-indigo-700 cursor-default'
                    : canManage
                      ? 'bg-gray-900 hover:bg-gray-700 text-white'
                      : 'bg-gray-100 text-gray-400 cursor-not-allowed'
                }`}
              >
                {isCurrent
                  ? 'Current plan'
                  : isPending
                    ? 'Opening checkout...'
                    : canManage
                      ? hasSubscription ? 'Switch to this' : 'Subscribe'
                      : 'Admin only'}
              </button>
            </div>
          )
        })}
      </div>

      {!canManage && (
        <p className="text-xs text-gray-500 mt-3">
          Only organization admins can change the plan.
        </p>
      )}

      {canManage && hasSubscription && (
        <div className="mt-4 pt-4 border-t border-gray-100 space-y-3">
          <button
            onClick={handleManage}
            disabled={pending !== null}
            className="block text-xs font-medium text-indigo-600 hover:text-indigo-700"
          >
            {pending === 'portal' ? 'Opening...' : 'Manage billing & invoices'}
          </button>

          {scheduledToCancel ? (
            <div className="rounded-lg bg-amber-50 border border-amber-200 p-3">
              <p className="text-xs text-amber-900">
                {isTrialing
                  ? `Your trial is set to cancel${endsAtLabel ? ` on ${endsAtLabel}` : ''}. You won't be charged.`
                  : `Your plan will cancel${endsAtLabel ? ` on ${endsAtLabel}` : ' at the end of the period'}. You keep access until then.`}
              </p>
              <button
                onClick={() => setCancellation(true)}
                disabled={pending !== null}
                className="mt-2 text-xs font-medium bg-amber-900 text-white px-3 py-1.5 rounded-md hover:bg-amber-800 disabled:opacity-50"
              >
                {pending === 'cancel' ? 'Working...' : 'Resume plan'}
              </button>
            </div>
          ) : confirmingCancel ? (
            <div className="rounded-lg bg-gray-50 border border-gray-200 p-3">
              <p className="text-xs text-gray-700">
                {isTrialing
                  ? `Cancel your trial? You'll keep access${endsAtLabel ? ` until ${endsAtLabel}` : ' until it ends'} and won't be charged.`
                  : `Cancel your plan? You'll keep access${endsAtLabel ? ` until ${endsAtLabel}` : ' until the end of your billing period'}, then it won't renew.`}
              </p>
              <div className="flex items-center gap-2 mt-2">
                <button
                  onClick={() => setCancellation(false)}
                  disabled={pending !== null}
                  className="text-xs font-medium bg-red-600 text-white px-3 py-1.5 rounded-md hover:bg-red-700 disabled:opacity-50"
                >
                  {pending === 'cancel' ? 'Cancelling...' : 'Yes, cancel'}
                </button>
                <button
                  onClick={() => setConfirmingCancel(false)}
                  disabled={pending !== null}
                  className="text-xs font-medium text-gray-600 hover:text-gray-800"
                >
                  Keep plan
                </button>
              </div>
            </div>
          ) : (
            <button
              onClick={() => setConfirmingCancel(true)}
              disabled={pending !== null}
              className="block text-xs font-medium text-gray-500 hover:text-red-600"
            >
              {isTrialing ? 'Cancel trial' : 'Cancel plan'}
            </button>
          )}
        </div>
      )}
    </section>
  )
}

import Link from 'next/link'
import { Sparkles, AlertCircle } from 'lucide-react'

export function TrialBanner({ daysLeft, expired, pastDue }: { daysLeft: number; expired: boolean; pastDue?: boolean }) {
  if (pastDue) {
    return (
      <div className="bg-red-50 border-b border-red-200 px-6 py-2.5 flex items-center justify-between gap-3">
        <div className="flex items-center gap-2 text-sm text-red-900">
          <AlertCircle className="w-4 h-4 shrink-0" />
          <span>Payment failed. Update your payment method to keep access.</span>
        </div>
        <Link
          href="/dashboard/settings"
          className="text-xs font-medium bg-red-700 text-white px-3 py-1.5 rounded-md hover:bg-red-800 transition-colors shrink-0"
        >
          Update card
        </Link>
      </div>
    )
  }

  if (expired) {
    return (
      <div className="bg-amber-50 border-b border-amber-200 px-6 py-2.5 flex items-center justify-between gap-3">
        <div className="flex items-center gap-2 text-sm text-amber-900">
          <AlertCircle className="w-4 h-4 shrink-0" />
          <span>Your free trial has ended. You&apos;re now on Starter limits.</span>
        </div>
        <Link
          href="/dashboard/settings"
          className="text-xs font-medium bg-amber-900 text-white px-3 py-1.5 rounded-md hover:bg-amber-800 transition-colors shrink-0"
        >
          Pick a plan
        </Link>
      </div>
    )
  }

  const urgent = daysLeft <= 3
  return (
    <div
      className={`border-b px-6 py-2.5 flex items-center justify-between gap-3 ${
        urgent ? 'bg-amber-50 border-amber-200' : 'bg-indigo-50 border-indigo-200'
      }`}
    >
      <div className={`flex items-center gap-2 text-sm ${urgent ? 'text-amber-900' : 'text-indigo-900'}`}>
        <Sparkles className="w-4 h-4 shrink-0" />
        <span>
          {daysLeft} {daysLeft === 1 ? 'day' : 'days'} left in your free trial — your card will be charged when it ends.
        </span>
      </div>
      <Link
        href="/dashboard/settings"
        className={`text-xs font-medium px-3 py-1.5 rounded-md transition-colors shrink-0 ${
          urgent
            ? 'bg-amber-900 text-white hover:bg-amber-800'
            : 'bg-indigo-600 text-white hover:bg-indigo-700'
        }`}
      >
        Manage billing
      </Link>
    </div>
  )
}

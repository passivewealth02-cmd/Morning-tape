import { sql } from '@/lib/db'
import { MaintenaLogo } from '@/components/brand/logo'
import { Check, Clock, Wrench, CircleDot, PauseCircle, XCircle } from 'lucide-react'

interface Props {
  params: Promise<{ id: string }>
}

export const metadata = {
  title: 'Track your maintenance request',
  robots: { index: false, follow: false },
}

type TrackRow = {
  id: string
  title: string
  status: string
  urgency: string
  created_at: string
  updated_at: string
  completed_at: string | null
  property_name: string | null
  unit_number: string | null
  vendor_name: string | null
}

const STEPS = [
  { key: 'new', label: 'Received', icon: CircleDot },
  { key: 'assigned', label: 'Assigned', icon: Clock },
  { key: 'in_progress', label: 'In progress', icon: Wrench },
  { key: 'completed', label: 'Completed', icon: Check },
]

const ORDER: Record<string, number> = { new: 0, assigned: 1, in_progress: 2, completed: 3 }

function formatDate(iso: string): string {
  try {
    return new Date(iso).toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: 'numeric' })
  } catch {
    return ''
  }
}

export default async function TrackPage({ params }: Props) {
  const { id } = await params

  // UUIDs are unguessable, so the ticket id itself acts as the access token.
  const isUuid = /^[0-9a-f-]{36}$/i.test(id)
  const rows = isUuid
    ? ((await sql`
        SELECT
          t.id, t.title, t.status, t.urgency, t.created_at, t.updated_at, t.completed_at,
          p.name AS property_name,
          u.unit_number,
          v.name AS vendor_name
        FROM maintenance_tickets t
        LEFT JOIN properties p ON p.id = t.property_id
        LEFT JOIN units u ON u.id = t.unit_id
        LEFT JOIN vendors v ON v.id = t.assigned_vendor_id
        WHERE t.id = ${id}
        LIMIT 1
      `) as unknown as TrackRow[])
    : []

  const ticket = rows[0]

  if (!ticket) {
    return (
      <div className="min-h-screen flex items-center justify-center px-6 bg-gray-50">
        <div className="max-w-sm w-full text-center">
          <h1 className="text-lg font-semibold text-gray-900">Request not found</h1>
          <p className="text-sm text-gray-500 mt-2">
            We couldn’t find this request. Please double-check the link from your confirmation.
          </p>
        </div>
      </div>
    )
  }

  const cancelled = ticket.status === 'cancelled'
  const waiting = ticket.status === 'waiting'
  const currentStep = ORDER[ticket.status] ?? 0
  const location = [ticket.property_name, ticket.unit_number && `Unit ${ticket.unit_number}`].filter(Boolean).join(' · ')

  return (
    <div className="min-h-screen bg-gray-50">
      <div className="max-w-lg mx-auto px-5 py-8 sm:py-10">
        <div className="flex items-center justify-center mb-6">
          <MaintenaLogo />
        </div>

        <div className="bg-white rounded-2xl border border-gray-200 p-6 shadow-sm">
          <p className="text-xs font-medium text-gray-400 uppercase tracking-wide">Maintenance request</p>
          <h1 className="text-xl font-semibold text-gray-900 mt-1">{ticket.title}</h1>
          {location && <p className="text-sm text-gray-500 mt-1">{location}</p>}
          <p className="text-xs text-gray-400 mt-1">Submitted {formatDate(ticket.created_at)}</p>

          {cancelled ? (
            <div className="mt-6 flex items-center gap-2 rounded-xl bg-gray-50 border border-gray-200 px-4 py-3">
              <XCircle className="w-5 h-5 text-gray-400" />
              <span className="text-sm font-medium text-gray-700">This request was cancelled.</span>
            </div>
          ) : (
            <>
              {/* Progress stepper */}
              <ol className="mt-6 space-y-0">
                {STEPS.map((step, i) => {
                  const done = i < currentStep || ticket.status === 'completed'
                  const active = i === currentStep && ticket.status !== 'completed'
                  const Icon = step.icon
                  return (
                    <li key={step.key} className="flex gap-3">
                      <div className="flex flex-col items-center">
                        <div
                          className={`w-8 h-8 rounded-full flex items-center justify-center shrink-0 ${
                            done ? 'bg-green-500 text-white' : active ? 'bg-indigo-600 text-white' : 'bg-gray-100 text-gray-400'
                          }`}
                        >
                          <Icon className="w-4 h-4" />
                        </div>
                        {i < STEPS.length - 1 && (
                          <div className={`w-0.5 flex-1 min-h-[24px] ${done ? 'bg-green-300' : 'bg-gray-100'}`} />
                        )}
                      </div>
                      <div className="pb-6">
                        <p className={`text-sm font-medium ${done || active ? 'text-gray-900' : 'text-gray-400'}`}>{step.label}</p>
                        {active && step.key === 'assigned' && ticket.vendor_name && (
                          <p className="text-xs text-gray-500 mt-0.5">Assigned to {ticket.vendor_name}</p>
                        )}
                        {active && <p className="text-xs text-indigo-600 mt-0.5">Current status</p>}
                      </div>
                    </li>
                  )
                })}
              </ol>

              {waiting && (
                <div className="mt-1 flex items-center gap-2 rounded-xl bg-amber-50 border border-amber-100 px-4 py-3">
                  <PauseCircle className="w-5 h-5 text-amber-500" />
                  <span className="text-sm font-medium text-amber-800">On hold — your manager will follow up.</span>
                </div>
              )}

              {ticket.vendor_name && ticket.status !== 'completed' && !waiting && (
                <p className="mt-1 text-xs text-gray-500">Handled by {ticket.vendor_name}</p>
              )}
              {ticket.status === 'completed' && (
                <p className="mt-1 text-sm font-medium text-green-700">
                  Completed{ticket.completed_at ? ` on ${formatDate(ticket.completed_at)}` : ''} 🎉
                </p>
              )}
            </>
          )}
        </div>

        <p className="text-center text-[11px] text-gray-400 mt-4">
          Last updated {formatDate(ticket.updated_at)} · Powered by Maintena
        </p>
      </div>
    </div>
  )
}

import { getSession } from '@/lib/auth'
import { sql } from '@/lib/db'
import { TicketsTable } from '@/components/tickets/tickets-table'
import Link from 'next/link'
import { Plus, FileEdit } from 'lucide-react'
import type { MaintenanceTicket } from '@/lib/db'

export default async function TicketsPage() {
  const session = await getSession()
  if (!session) return null

  const orgId = session.user.organization_id!

  const allTickets = await sql`
    SELECT
      t.*,
      p.name as property_name,
      p.address as property_address,
      u.unit_number,
      v.name as vendor_name,
      v.trade_type as vendor_trade_type
    FROM maintenance_tickets t
    LEFT JOIN properties p ON t.property_id = p.id
    LEFT JOIN units u ON t.unit_id = u.id
    LEFT JOIN vendors v ON t.assigned_vendor_id = v.id
    WHERE t.organization_id = ${orgId}
    ORDER BY t.created_at DESC
    LIMIT 500
  ` as MaintenanceTicket[]

  const drafts = allTickets.filter(t => t.is_draft)
  const tickets = allTickets.filter(t => !t.is_draft)

  return (
    <div className="p-6">
      <div className="flex items-center justify-between mb-6">
        <div>
          <h1 className="text-xl font-semibold text-gray-900">Tickets</h1>
          <p className="text-sm text-gray-500 mt-0.5">{tickets.length} active{drafts.length > 0 ? ` · ${drafts.length} draft${drafts.length !== 1 ? 's' : ''}` : ''}</p>
        </div>
        <Link
          href="/dashboard/tickets/new"
          className="inline-flex items-center gap-1.5 bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-medium px-4 py-2 rounded-lg transition-colors"
        >
          <Plus className="w-4 h-4" />
          New Ticket
        </Link>
      </div>

      {drafts.length > 0 && (
        <div className="mb-6 rounded-xl border border-amber-200 bg-amber-50/60 p-4">
          <div className="flex items-center gap-2 mb-3">
            <FileEdit className="w-4 h-4 text-amber-600" />
            <h2 className="text-sm font-semibold text-amber-900">Drafts ({drafts.length})</h2>
            <span className="text-xs text-amber-700">Not yet triaged or dispatched</span>
          </div>
          <ul className="space-y-2">
            {drafts.map(d => (
              <li key={d.id}>
                <Link
                  href={`/dashboard/tickets/${d.id}`}
                  className="flex items-center justify-between gap-3 rounded-lg bg-white border border-amber-100 px-4 py-2.5 hover:border-amber-300 transition-colors"
                >
                  <div className="min-w-0">
                    <p className="text-sm font-medium text-gray-900 truncate">{d.title || 'Untitled draft'}</p>
                    <p className="text-xs text-gray-400 truncate">
                      {d.property_name ? d.property_name : 'No property'}
                      {d.unit_number ? ` · Unit ${d.unit_number}` : ''}
                    </p>
                  </div>
                  <span className="text-xs font-medium text-amber-700 shrink-0">Open →</span>
                </Link>
              </li>
            ))}
          </ul>
        </div>
      )}

      <TicketsTable tickets={tickets} />
    </div>
  )
}

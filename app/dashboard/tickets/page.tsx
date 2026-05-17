import { getSession } from '@/lib/auth'
import { sql } from '@/lib/db'
import { TicketsTable } from '@/components/tickets/tickets-table'
import Link from 'next/link'
import { Plus } from 'lucide-react'
import type { MaintenanceTicket } from '@/lib/db'

export default async function TicketsPage() {
  const session = await getSession()
  if (!session) return null

  const orgId = session.user.organization_id!

  const tickets = await sql`
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

  return (
    <div className="p-6">
      <div className="flex items-center justify-between mb-6">
        <div>
          <h1 className="text-xl font-semibold text-gray-900">Tickets</h1>
          <p className="text-sm text-gray-500 mt-0.5">{tickets.length} total</p>
        </div>
        <Link
          href="/dashboard/tickets/new"
          className="inline-flex items-center gap-1.5 bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-medium px-4 py-2 rounded-lg transition-colors"
        >
          <Plus className="w-4 h-4" />
          New Ticket
        </Link>
      </div>

      <TicketsTable tickets={tickets} />
    </div>
  )
}

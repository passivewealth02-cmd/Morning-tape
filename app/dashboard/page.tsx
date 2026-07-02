import { getSession } from '@/lib/auth'
import { sql } from '@/lib/db'
import { KanbanBoard } from '@/components/tickets/kanban-board'
import Link from 'next/link'
import { Plus } from 'lucide-react'
import type { MaintenanceTicket } from '@/lib/db'

export default async function DashboardPage() {
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
      AND t.status != 'cancelled'
      AND t.is_draft = FALSE
    ORDER BY
      CASE t.urgency
        WHEN 'emergency' THEN 1
        WHEN 'high' THEN 2
        WHEN 'medium' THEN 3
        WHEN 'low' THEN 4
      END,
      t.created_at DESC
    LIMIT 200
  ` as MaintenanceTicket[]

  const stats = {
    total: tickets.length,
    new: tickets.filter(t => t.status === 'new').length,
    inProgress: tickets.filter(t => t.status === 'in_progress' || t.status === 'assigned').length,
    overdue: tickets.filter(t => {
      if (!t.sla_due_at) return false
      return new Date(t.sla_due_at) < new Date() && t.status !== 'completed'
    }).length,
  }

  return (
    <div className="p-6">
      {/* Header */}
      <div className="flex items-center justify-between mb-6">
        <div>
          <h1 className="text-xl font-semibold text-gray-900">Dashboard</h1>
          <p className="text-sm text-gray-500 mt-0.5">
            {stats.total} active ticket{stats.total !== 1 ? 's' : ''}
          </p>
        </div>
        <Link
          href="/dashboard/tickets/new"
          className="inline-flex items-center gap-1.5 bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-medium px-4 py-2 rounded-lg transition-colors"
        >
          <Plus className="w-4 h-4" />
          New Ticket
        </Link>
      </div>

      {/* Stats */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 sm:gap-4 mb-6">
        {[
          { label: 'Total Open', value: stats.total, color: 'text-gray-900' },
          { label: 'New / Unassigned', value: stats.new, color: 'text-blue-700' },
          { label: 'In Progress', value: stats.inProgress, color: 'text-yellow-700' },
          { label: 'Overdue', value: stats.overdue, color: 'text-red-600' },
        ].map(stat => (
          <div key={stat.label} className="bg-white rounded-lg border border-gray-200 p-4">
            <p className="text-xs text-gray-500 mb-1">{stat.label}</p>
            <p className={`text-2xl font-semibold ${stat.color}`}>{stat.value}</p>
          </div>
        ))}
      </div>

      {/* Kanban Board */}
      <KanbanBoard tickets={tickets} />
    </div>
  )
}

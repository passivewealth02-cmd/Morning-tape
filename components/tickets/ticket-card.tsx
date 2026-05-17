import Link from 'next/link'
import type { MaintenanceTicket } from '@/lib/db'
import { UrgencyBadge } from './urgency-badge'
import { formatDistanceToNow } from 'date-fns'

interface TicketCardProps {
  ticket: MaintenanceTicket
  draggable?: boolean
  onDragStart?: (e: React.DragEvent) => void
}

export function TicketCard({ ticket, draggable, onDragStart }: TicketCardProps) {
  const isOverdue =
    ticket.sla_due_at &&
    new Date(ticket.sla_due_at) < new Date() &&
    ticket.status !== 'completed'

  return (
    <Link href={`/dashboard/tickets/${ticket.id}`}>
      <div
        draggable={draggable}
        onDragStart={onDragStart}
        className={`bg-white rounded-lg border p-3 cursor-pointer hover:shadow-sm transition-shadow ${
          isOverdue ? 'border-red-200' : 'border-gray-200'
        }`}
      >
        <div className="flex items-start justify-between gap-2 mb-2">
          <p className="text-sm font-medium text-gray-900 leading-tight line-clamp-2">
            {ticket.title}
          </p>
          <UrgencyBadge urgency={ticket.urgency} />
        </div>

        {ticket.property_name && (
          <p className="text-xs text-gray-500 mb-1">
            {ticket.property_name}
            {ticket.unit_number && ` · Unit ${ticket.unit_number}`}
          </p>
        )}

        {ticket.ai_category && (
          <p className="text-xs text-indigo-600 mb-2 capitalize">
            {ticket.ai_category.replace('_', ' ')}
          </p>
        )}

        <div className="flex items-center justify-between mt-2">
          {ticket.vendor_name ? (
            <span className="text-xs text-gray-500 truncate">{ticket.vendor_name}</span>
          ) : (
            <span className="text-xs text-gray-400">Unassigned</span>
          )}
          <span className="text-xs text-gray-400 shrink-0 ml-2">
            {formatDistanceToNow(new Date(ticket.created_at), { addSuffix: true })}
          </span>
        </div>

        {isOverdue && (
          <div className="mt-2 text-xs text-red-600 font-medium">Overdue</div>
        )}
      </div>
    </Link>
  )
}

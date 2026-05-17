import type { TicketUrgency } from '@/lib/db'

const config: Record<TicketUrgency, { label: string; className: string }> = {
  emergency: { label: 'Emergency', className: 'bg-red-100 text-red-700' },
  high: { label: 'High', className: 'bg-orange-100 text-orange-700' },
  medium: { label: 'Medium', className: 'bg-yellow-100 text-yellow-700' },
  low: { label: 'Low', className: 'bg-green-100 text-green-700' },
}

export function UrgencyBadge({ urgency }: { urgency: TicketUrgency }) {
  const { label, className } = config[urgency] ?? config.medium
  return (
    <span className={`inline-flex items-center px-1.5 py-0.5 rounded text-xs font-medium shrink-0 ${className}`}>
      {label}
    </span>
  )
}

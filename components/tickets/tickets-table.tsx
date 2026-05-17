'use client'

import { useState } from 'react'
import Link from 'next/link'
import type { MaintenanceTicket, TicketStatus, TicketUrgency } from '@/lib/db'
import { UrgencyBadge } from './urgency-badge'
import { formatDistanceToNow } from 'date-fns'

const STATUS_LABELS: Record<TicketStatus, string> = {
  new: 'New',
  assigned: 'Assigned',
  in_progress: 'In Progress',
  waiting: 'Waiting',
  completed: 'Completed',
  cancelled: 'Cancelled',
}

const STATUS_COLORS: Record<TicketStatus, string> = {
  new: 'bg-blue-50 text-blue-700',
  assigned: 'bg-purple-50 text-purple-700',
  in_progress: 'bg-yellow-50 text-yellow-700',
  waiting: 'bg-gray-100 text-gray-600',
  completed: 'bg-green-50 text-green-700',
  cancelled: 'bg-gray-100 text-gray-500',
}

interface TicketsTableProps {
  tickets: MaintenanceTicket[]
}

export function TicketsTable({ tickets }: TicketsTableProps) {
  const [statusFilter, setStatusFilter] = useState<string>('all')
  const [urgencyFilter, setUrgencyFilter] = useState<string>('all')
  const [search, setSearch] = useState('')

  const filtered = tickets.filter(t => {
    if (statusFilter !== 'all' && t.status !== statusFilter) return false
    if (urgencyFilter !== 'all' && t.urgency !== urgencyFilter) return false
    if (search && !t.title.toLowerCase().includes(search.toLowerCase())) return false
    return true
  })

  return (
    <div className="space-y-4">
      {/* Filters */}
      <div className="flex items-center gap-3 flex-wrap">
        <input
          type="text"
          placeholder="Search tickets..."
          value={search}
          onChange={e => setSearch(e.target.value)}
          className="border border-gray-200 rounded-lg px-3 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-white"
        />

        <select
          value={statusFilter}
          onChange={e => setStatusFilter(e.target.value)}
          className="border border-gray-200 rounded-lg px-3 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-white"
        >
          <option value="all">All statuses</option>
          {Object.entries(STATUS_LABELS).map(([v, l]) => (
            <option key={v} value={v}>{l}</option>
          ))}
        </select>

        <select
          value={urgencyFilter}
          onChange={e => setUrgencyFilter(e.target.value)}
          className="border border-gray-200 rounded-lg px-3 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-white"
        >
          <option value="all">All urgencies</option>
          {(['emergency', 'high', 'medium', 'low'] as TicketUrgency[]).map(u => (
            <option key={u} value={u}>{u.charAt(0).toUpperCase() + u.slice(1)}</option>
          ))}
        </select>

        {(statusFilter !== 'all' || urgencyFilter !== 'all' || search) && (
          <button
            onClick={() => { setStatusFilter('all'); setUrgencyFilter('all'); setSearch('') }}
            className="text-sm text-gray-500 hover:text-gray-700"
          >
            Clear filters
          </button>
        )}

        <span className="text-sm text-gray-400 ml-auto">{filtered.length} results</span>
      </div>

      {/* Table */}
      <div className="bg-white rounded-lg border border-gray-200 overflow-hidden">
        <table className="w-full text-sm">
          <thead>
            <tr className="border-b border-gray-100 bg-gray-50">
              <th className="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Title</th>
              <th className="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Property</th>
              <th className="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Status</th>
              <th className="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Urgency</th>
              <th className="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Vendor</th>
              <th className="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Created</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-50">
            {filtered.length === 0 ? (
              <tr>
                <td colSpan={6} className="px-4 py-8 text-center text-gray-400 text-sm">
                  No tickets found
                </td>
              </tr>
            ) : (
              filtered.map(ticket => (
                <tr key={ticket.id} className="hover:bg-gray-50 transition-colors">
                  <td className="px-4 py-3">
                    <Link href={`/dashboard/tickets/${ticket.id}`} className="font-medium text-gray-900 hover:text-indigo-600 line-clamp-1">
                      {ticket.title}
                    </Link>
                    {ticket.ai_category && (
                      <p className="text-xs text-gray-400 capitalize mt-0.5">{ticket.ai_category.replace('_', ' ')}</p>
                    )}
                  </td>
                  <td className="px-4 py-3 text-gray-500">
                    {ticket.property_name || '—'}
                    {ticket.unit_number && <span className="text-gray-400"> · {ticket.unit_number}</span>}
                  </td>
                  <td className="px-4 py-3">
                    <span className={`inline-flex items-center px-2 py-0.5 rounded text-xs font-medium ${STATUS_COLORS[ticket.status]}`}>
                      {STATUS_LABELS[ticket.status]}
                    </span>
                  </td>
                  <td className="px-4 py-3">
                    <UrgencyBadge urgency={ticket.urgency} />
                  </td>
                  <td className="px-4 py-3 text-gray-500">
                    {ticket.vendor_name || <span className="text-gray-300">Unassigned</span>}
                  </td>
                  <td className="px-4 py-3 text-gray-400 text-xs">
                    {formatDistanceToNow(new Date(ticket.created_at), { addSuffix: true })}
                  </td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
    </div>
  )
}

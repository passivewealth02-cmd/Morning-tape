'use client'

import { useState } from 'react'
import type { MaintenanceTicket, TicketStatus } from '@/lib/db'
import { TicketCard } from './ticket-card'

const COLUMNS: { id: TicketStatus; label: string }[] = [
  { id: 'new', label: 'New' },
  { id: 'assigned', label: 'Assigned' },
  { id: 'in_progress', label: 'In Progress' },
  { id: 'waiting', label: 'Waiting' },
  { id: 'completed', label: 'Completed' },
]

interface KanbanBoardProps {
  tickets: MaintenanceTicket[]
}

export function KanbanBoard({ tickets: initial }: KanbanBoardProps) {
  const [tickets, setTickets] = useState(initial)

  const byStatus = (status: TicketStatus) => tickets.filter(t => t.status === status)

  const moveTicket = async (ticketId: string, newStatus: TicketStatus) => {
    setTickets(prev =>
      prev.map(t => t.id === ticketId ? { ...t, status: newStatus } : t)
    )
    await fetch(`/api/tickets/${ticketId}`, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ status: newStatus }),
    })
  }

  return (
    <div className="flex gap-4 overflow-x-auto pb-4">
      {COLUMNS.map(col => {
        const colTickets = byStatus(col.id)
        return (
          <div
            key={col.id}
            className="flex-shrink-0 w-64"
            onDragOver={e => e.preventDefault()}
            onDrop={e => {
              const id = e.dataTransfer.getData('ticketId')
              if (id) moveTicket(id, col.id)
            }}
          >
            <div className="flex items-center gap-2 mb-3">
              <span className="text-xs font-semibold text-gray-600 uppercase tracking-wide">
                {col.label}
              </span>
              <span className="text-xs bg-gray-100 text-gray-600 px-1.5 py-0.5 rounded-md tabular-nums">
                {colTickets.length}
              </span>
            </div>

            <div className="space-y-2 min-h-[120px]">
              {colTickets.map(ticket => (
                <TicketCard
                  key={ticket.id}
                  ticket={ticket}
                  draggable
                  onDragStart={e => e.dataTransfer.setData('ticketId', ticket.id)}
                />
              ))}
              {colTickets.length === 0 && (
                <div className="border-2 border-dashed border-gray-200 rounded-lg h-16 flex items-center justify-center">
                  <span className="text-xs text-gray-400">Drop here</span>
                </div>
              )}
            </div>
          </div>
        )
      })}
    </div>
  )
}

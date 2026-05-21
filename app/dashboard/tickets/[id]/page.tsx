import { getSession } from '@/lib/auth'
import { sql } from '@/lib/db'
import { notFound } from 'next/navigation'
import { TicketDetail } from '@/components/tickets/ticket-detail'
import type { MaintenanceTicket, TicketMessage, ActivityLog, Vendor, TicketFile } from '@/lib/db'

interface Props {
  params: Promise<{ id: string }>
}

export default async function TicketDetailPage({ params }: Props) {
  const { id } = await params
  const session = await getSession()
  if (!session) return null

  const orgId = session.user.organization_id!

  const [ticketRows, messages, logs, vendors, files] = await Promise.all([
    sql`
      SELECT
        t.*,
        p.name as property_name,
        p.address as property_address,
        u.unit_number,
        v.name as vendor_name,
        v.trade_type as vendor_trade_type,
        v.phone as vendor_phone,
        v.email as vendor_email
      FROM maintenance_tickets t
      LEFT JOIN properties p ON t.property_id = p.id
      LEFT JOIN units u ON t.unit_id = u.id
      LEFT JOIN vendors v ON t.assigned_vendor_id = v.id
      WHERE t.id = ${id} AND t.organization_id = ${orgId}
    `,
    sql`
      SELECT m.*, u.name as sender_name
      FROM ticket_messages m
      LEFT JOIN users u ON m.sender_id = u.id
      WHERE m.ticket_id = ${id}
      ORDER BY m.created_at ASC
    `,
    sql`
      SELECT a.*, u.name as user_name
      FROM activity_logs a
      LEFT JOIN users u ON a.user_id = u.id
      WHERE a.ticket_id = ${id}
      ORDER BY a.created_at DESC
    `,
    sql`
      SELECT id, name, trade_type, availability, rating, phone, email
      FROM vendors
      WHERE organization_id = ${orgId}
      ORDER BY rating DESC, name ASC
    `,
    sql`
      SELECT * FROM ticket_files
      WHERE ticket_id = ${id}
      ORDER BY created_at ASC
    `,
  ])

  if (ticketRows.length === 0) notFound()

  return (
    <div className="p-6">
      <TicketDetail
        ticket={ticketRows[0] as unknown as MaintenanceTicket}
        messages={messages as unknown as TicketMessage[]}
        activityLogs={logs as unknown as ActivityLog[]}
        vendors={vendors as unknown as Vendor[]}
        files={files as unknown as TicketFile[]}
        currentUserId={session.user.id}
      />
    </div>
  )
}

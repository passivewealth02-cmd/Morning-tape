import { NextRequest, NextResponse } from 'next/server'
import { getSession, logActivity } from '@/lib/auth'
import { sql, type MaintenanceTicket } from '@/lib/db'
import { sendTenantVendorAssignedEmail, sendTenantStatusUpdateEmail } from '@/lib/email'

export async function GET(
  _request: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const session = await getSession()
    if (!session) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
    }

    const { user } = session
    if (!user.organization_id) {
      return NextResponse.json({ error: 'No organization' }, { status: 403 })
    }

    const { id } = await params

    const tickets = await sql`
      SELECT
        t.*,
        p.name AS property_name,
        p.address AS property_address,
        u.unit_number,
        v.name AS vendor_name,
        v.trade_type AS vendor_trade_type
      FROM maintenance_tickets t
      LEFT JOIN properties p ON p.id = t.property_id
      LEFT JOIN units u ON u.id = t.unit_id
      LEFT JOIN vendors v ON v.id = t.assigned_vendor_id
      WHERE t.id = ${id}
        AND t.organization_id = ${user.organization_id}
    `

    if (tickets.length === 0) {
      return NextResponse.json({ error: 'Ticket not found' }, { status: 404 })
    }

    const ticket = tickets[0]

    const messages = await sql`
      SELECT m.*, u.name AS sender_name
      FROM ticket_messages m
      LEFT JOIN users u ON u.id = m.sender_id
      WHERE m.ticket_id = ${id}
      ORDER BY m.created_at ASC
    `

    const files = await sql`
      SELECT * FROM ticket_files
      WHERE ticket_id = ${id}
      ORDER BY created_at ASC
    `

    const activityLogs = await sql`
      SELECT a.*, u.name AS user_name
      FROM activity_logs a
      LEFT JOIN users u ON u.id = a.user_id
      WHERE a.ticket_id = ${id}
      ORDER BY a.created_at DESC
    `

    return NextResponse.json({ ...ticket, messages, files, activity_logs: activityLogs })
  } catch (error) {
    console.error('Error fetching ticket:', error)
    return NextResponse.json({ error: 'Internal server error' }, { status: 500 })
  }
}

export async function PATCH(
  request: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const session = await getSession()
    if (!session) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
    }

    const { user } = session
    if (!user.organization_id) {
      return NextResponse.json({ error: 'No organization' }, { status: 403 })
    }

    const { id } = await params

    const existing = (await sql`
      SELECT * FROM maintenance_tickets
      WHERE id = ${id} AND organization_id = ${user.organization_id}
    `) as unknown as MaintenanceTicket[]

    if (existing.length === 0) {
      return NextResponse.json({ error: 'Ticket not found' }, { status: 404 })
    }

    const current = existing[0]
    const body = await request.json()
    const {
      status,
      urgency,
      assigned_vendor_id,
      title,
      description,
      tenant_name,
      tenant_email,
      tenant_phone,
    } = body

    const newStatus = status ?? current.status
    const newUrgency = urgency ?? current.urgency
    const newVendorId = assigned_vendor_id !== undefined ? assigned_vendor_id : current.assigned_vendor_id
    const assignedAt = assigned_vendor_id && !current.assigned_vendor_id ? new Date().toISOString() : current.assigned_at
    const completedAt = status === 'completed' && current.status !== 'completed' ? new Date().toISOString() : current.completed_at

    const updated = (await sql`
      UPDATE maintenance_tickets
      SET
        status = ${newStatus},
        urgency = ${newUrgency},
        assigned_vendor_id = ${newVendorId},
        assigned_at = ${assignedAt},
        completed_at = ${completedAt},
        title = ${title ?? current.title},
        description = ${description ?? current.description},
        tenant_name = ${tenant_name !== undefined ? tenant_name : current.tenant_name},
        tenant_email = ${tenant_email !== undefined ? tenant_email : current.tenant_email},
        tenant_phone = ${tenant_phone !== undefined ? tenant_phone : current.tenant_phone},
        updated_at = NOW()
      WHERE id = ${id}
        AND organization_id = ${user.organization_id}
      RETURNING *
    `) as unknown as MaintenanceTicket[]

    if (status && status !== current.status) {
      await logActivity({
        organizationId: user.organization_id,
        ticketId: id,
        userId: user.id,
        actionType: 'status_changed',
        description: `Status changed from ${current.status} to ${status}`,
        metadata: { from: current.status, to: status },
      })
    }

    const vendorChanged = assigned_vendor_id && assigned_vendor_id !== current.assigned_vendor_id

    if (vendorChanged) {
      await logActivity({
        organizationId: user.organization_id,
        ticketId: id,
        userId: user.id,
        actionType: 'vendor_assigned',
        description: `Vendor assigned to ticket`,
        metadata: { vendor_id: assigned_vendor_id },
      })

      // Notify tenant that a vendor has been assigned
      const tenantEmail = updated[0].tenant_email ?? current.tenant_email
      if (tenantEmail) {
        const vendorRows = (await sql`SELECT name FROM vendors WHERE id = ${assigned_vendor_id}`) as unknown as { name: string }[]
        const vendorName = vendorRows[0]?.name ?? 'a vendor'
        const propertyRows = current.property_id
          ? (await sql`SELECT name FROM properties WHERE id = ${current.property_id}`) as unknown as { name: string }[]
          : []
        await sendTenantVendorAssignedEmail(
          tenantEmail,
          updated[0].tenant_name ?? current.tenant_name,
          updated[0].title ?? current.title,
          vendorName,
          propertyRows[0]?.name ?? null,
        )
      }
    }

    // Notify tenant on meaningful status transitions (skip if we already sent a vendor-assigned email)
    const statusChanged = status && status !== current.status
    const notifiableStatus = ['in_progress', 'completed', 'waiting'] as const
    type NotifiableStatus = typeof notifiableStatus[number]
    if (!vendorChanged && statusChanged && (notifiableStatus as readonly string[]).includes(status)) {
      const tenantEmail = updated[0].tenant_email ?? current.tenant_email
      if (tenantEmail) {
        const propertyRows = current.property_id
          ? (await sql`SELECT name FROM properties WHERE id = ${current.property_id}`) as unknown as { name: string }[]
          : []
        await sendTenantStatusUpdateEmail(
          tenantEmail,
          updated[0].tenant_name ?? current.tenant_name,
          updated[0].title ?? current.title,
          status as NotifiableStatus,
          propertyRows[0]?.name ?? null,
        )
      }
    }

    return NextResponse.json(updated[0])
  } catch (error) {
    console.error('Error updating ticket:', error)
    return NextResponse.json({ error: 'Internal server error' }, { status: 500 })
  }
}

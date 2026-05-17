import { NextRequest, NextResponse } from 'next/server'
import { getSession, logActivity } from '@/lib/auth'
import { sql } from '@/lib/db'
import { analyzeMaintenanceTicket } from '@/lib/anthropic'

export async function GET(request: NextRequest) {
  try {
    const session = await getSession()
    if (!session) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
    }

    const { user } = session
    if (!user.organization_id) {
      return NextResponse.json({ error: 'No organization' }, { status: 403 })
    }

    const { searchParams } = new URL(request.url)
    const status = searchParams.get('status')
    const urgency = searchParams.get('urgency')
    const propertyId = searchParams.get('property_id')

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
      WHERE t.organization_id = ${user.organization_id}
        AND (${status}::text IS NULL OR t.status = ${status})
        AND (${urgency}::text IS NULL OR t.urgency = ${urgency})
        AND (${propertyId}::text IS NULL OR t.property_id = ${propertyId})
      ORDER BY t.created_at DESC
    `

    return NextResponse.json(tickets)
  } catch (error) {
    console.error('Error fetching tickets:', error)
    return NextResponse.json({ error: 'Internal server error' }, { status: 500 })
  }
}

export async function POST(request: NextRequest) {
  try {
    const session = await getSession()
    if (!session) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
    }

    const { user } = session
    if (!user.organization_id) {
      return NextResponse.json({ error: 'No organization' }, { status: 403 })
    }

    const body = await request.json()
    const {
      title,
      description,
      urgency = 'medium',
      property_id = null,
      unit_id = null,
      tenant_name = null,
      tenant_email = null,
      tenant_phone = null,
    } = body

    if (!title || !description) {
      return NextResponse.json({ error: 'Title and description are required' }, { status: 400 })
    }

    const inserted = await sql`
      INSERT INTO maintenance_tickets (
        organization_id, property_id, unit_id, title, description, urgency,
        status, tenant_name, tenant_email, tenant_phone, created_by
      )
      VALUES (
        ${user.organization_id}, ${property_id}, ${unit_id}, ${title}, ${description}, ${urgency},
        'new', ${tenant_name}, ${tenant_email}, ${tenant_phone}, ${user.id}
      )
      RETURNING *
    `

    const ticket = inserted[0]

    try {
      const analysis = await analyzeMaintenanceTicket(title, description)
      await sql`
        UPDATE maintenance_tickets
        SET
          ai_category = ${analysis.category},
          ai_vendor_type = ${analysis.vendor_type},
          ai_summary = ${analysis.summary},
          ai_escalation_risk = ${analysis.escalation_risk},
          urgency = ${analysis.urgency}
        WHERE id = ${ticket.id}
      `
      ticket.ai_category = analysis.category
      ticket.ai_vendor_type = analysis.vendor_type
      ticket.ai_summary = analysis.summary
      ticket.ai_escalation_risk = analysis.escalation_risk
      ticket.urgency = analysis.urgency
    } catch (aiError) {
      console.error('AI analysis failed:', aiError)
    }

    await logActivity({
      organizationId: user.organization_id,
      ticketId: ticket.id,
      userId: user.id,
      actionType: 'ticket_created',
      description: `Ticket created: ${title}`,
      metadata: { urgency: ticket.urgency },
    })

    return NextResponse.json(ticket, { status: 201 })
  } catch (error) {
    console.error('Error creating ticket:', error)
    return NextResponse.json({ error: 'Internal server error' }, { status: 500 })
  }
}

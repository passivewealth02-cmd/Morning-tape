import { NextRequest, NextResponse } from 'next/server'
import { getSession } from '@/lib/auth'
import { sql } from '@/lib/db'
import { createTicketWithAI } from '@/lib/tickets'

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

    const ticket = await createTicketWithAI({
      organization_id: user.organization_id,
      title,
      description,
      urgency,
      property_id,
      unit_id,
      tenant_name,
      tenant_email,
      tenant_phone,
      created_by: user.id,
      source: 'manual',
    })

    return NextResponse.json(ticket, { status: 201 })
  } catch (error) {
    console.error('Error creating ticket:', error)
    return NextResponse.json({ error: 'Internal server error' }, { status: 500 })
  }
}

import { NextRequest, NextResponse } from 'next/server'
import { getSession } from '@/lib/auth'
import { sql, type Organization } from '@/lib/db'
import { createTicketWithAI } from '@/lib/tickets'
import { checkResourceLimit, getEffectivePlan, getUsage } from '@/lib/plans'

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
      unit_id: unitIdInput = null,
      unit_number = null,
      tenant_name = null,
      tenant_email = null,
      tenant_phone = null,
    } = body

    if (!title || !description) {
      return NextResponse.json({ error: 'Title and description are required' }, { status: 400 })
    }

    const orgRows = (await sql`
      SELECT * FROM organizations WHERE id = ${user.organization_id}
    `) as unknown as Organization[]
    if (orgRows.length === 0) {
      return NextResponse.json({ error: 'Organization not found' }, { status: 404 })
    }
    const usage = await getUsage(user.organization_id)
    const check = checkResourceLimit(getEffectivePlan(orgRows[0]), 'tickets_per_month', usage.tickets_this_month)
    if (!check.allowed) {
      return NextResponse.json(
        {
          error: `You've hit your monthly ticket limit (${check.current}/${check.limit}). Upgrade to ${check.upgrade_to ?? 'a higher plan'} to keep creating tickets.`,
          limit_exceeded: true,
          upgrade_to: check.upgrade_to,
        },
        { status: 402 }
      )
    }

    let resolvedUnitId: string | null = unitIdInput
    if (!resolvedUnitId && unit_number && property_id) {
      const unitMatch = (await sql`
        SELECT id FROM units
        WHERE property_id = ${property_id}
          AND LOWER(unit_number) = LOWER(${unit_number})
        LIMIT 1
      `) as unknown as Array<{ id: string }>
      if (unitMatch.length > 0) resolvedUnitId = unitMatch[0].id
    }

    const ticket = await createTicketWithAI({
      organization_id: user.organization_id,
      title,
      description: unit_number && !resolvedUnitId
        ? `${description}\n\n(Unit: ${unit_number})`
        : description,
      urgency,
      property_id,
      unit_id: resolvedUnitId,
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

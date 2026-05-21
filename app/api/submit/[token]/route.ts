import { NextRequest, NextResponse } from 'next/server'
import { sql, type Organization, type Property } from '@/lib/db'
import { createTicketWithAI } from '@/lib/tickets'

const URGENCIES = ['low', 'medium', 'high', 'emergency'] as const
type Urgency = (typeof URGENCIES)[number]

function clamp(value: unknown, max: number): string {
  return typeof value === 'string' ? value.trim().slice(0, max) : ''
}

export async function POST(
  request: NextRequest,
  { params }: { params: Promise<{ token: string }> }
) {
  try {
    const { token } = await params

    const orgs = (await sql`
      SELECT * FROM organizations WHERE inbox_token = ${token} LIMIT 1
    `) as unknown as Organization[]

    if (orgs.length === 0) {
      return NextResponse.json({ error: 'This submission link is not valid.' }, { status: 404 })
    }
    const org = orgs[0]

    const body = (await request.json()) as Record<string, unknown>

    const title = clamp(body.title, 200)
    const description = clamp(body.description, 5000)
    const tenantName = clamp(body.tenant_name, 120) || null
    const tenantEmail = clamp(body.tenant_email, 200) || null
    const tenantPhone = clamp(body.tenant_phone, 50) || null
    const propertyHint = clamp(body.property_hint, 200)
    const unitHint = clamp(body.unit_number, 50)
    const rawUrgency = clamp(body.urgency, 20)
    const urgency: Urgency = (URGENCIES as readonly string[]).includes(rawUrgency)
      ? (rawUrgency as Urgency)
      : 'medium'

    if (!title || !description) {
      return NextResponse.json({ error: 'Please include a title and a description.' }, { status: 400 })
    }

    // Best-effort property/unit matching from the free-text hints (same approach as email intake)
    let propertyId: string | null = null
    let unitId: string | null = null
    if (propertyHint) {
      const hint = `%${propertyHint.toLowerCase()}%`
      const props = (await sql`
        SELECT * FROM properties
        WHERE organization_id = ${org.id}
          AND (LOWER(name) LIKE ${hint} OR LOWER(address) LIKE ${hint})
        LIMIT 1
      `) as unknown as Property[]
      if (props.length > 0) {
        propertyId = props[0].id
        if (unitHint) {
          const unitMatch = (await sql`
            SELECT id FROM units
            WHERE property_id = ${propertyId}
              AND LOWER(unit_number) = LOWER(${unitHint})
            LIMIT 1
          `) as unknown as Array<{ id: string }>
          if (unitMatch.length > 0) unitId = unitMatch[0].id
        }
      }
    }

    // Keep the tenant-provided location in the description when we couldn't match a property
    const locationNote =
      !propertyId && (propertyHint || unitHint)
        ? `\n\nLocation provided by tenant: ${[propertyHint, unitHint && `Unit ${unitHint}`].filter(Boolean).join(', ')}`
        : ''

    const ticket = await createTicketWithAI({
      organization_id: org.id,
      title,
      description: description + locationNote,
      urgency,
      property_id: propertyId,
      unit_id: unitId,
      tenant_name: tenantName,
      tenant_email: tenantEmail,
      tenant_phone: tenantPhone,
      source: 'web',
    })

    return NextResponse.json({ success: true, ticket_id: ticket.id }, { status: 201 })
  } catch (error) {
    console.error('Tenant submission error:', error)
    return NextResponse.json({ error: 'Something went wrong. Please try again.' }, { status: 500 })
  }
}

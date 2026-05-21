import { NextRequest, NextResponse } from 'next/server'
import { sql, type Organization, type Property } from '@/lib/db'
import { extractTicketFromEmail } from '@/lib/anthropic'
import { createTicketWithAI } from '@/lib/tickets'
import { rateLimit, tooManyRequests } from '@/lib/rate-limit'

type EmailPayload = {
  from?: string
  subject?: string
  text?: string
  html?: string
  body?: string
}

function normalizePayload(raw: Record<string, unknown>): EmailPayload {
  return {
    from: pickString(raw, ['from', 'From', 'sender', 'fromEmail']),
    subject: pickString(raw, ['subject', 'Subject']),
    text: pickString(raw, ['text', 'plain', 'body-plain', 'bodyPlain']),
    html: pickString(raw, ['html', 'body-html', 'bodyHtml']),
    body: pickString(raw, ['body', 'message']),
  }
}

function pickString(obj: Record<string, unknown>, keys: string[]): string | undefined {
  for (const k of keys) {
    const v = obj[k]
    if (typeof v === 'string' && v.trim().length > 0) return v
  }
  return undefined
}

function stripHtml(html: string): string {
  return html
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/\s+/g, ' ')
    .trim()
}

function parseFromAddress(from: string): string {
  const match = from.match(/<([^>]+)>/)
  return (match ? match[1] : from).trim()
}

export async function POST(
  request: NextRequest,
  { params }: { params: Promise<{ token: string }> }
) {
  try {
    const { token } = await params

    // Bound inbound volume per inbox to cap AI cost and ticket spam.
    const inboundLimit = await rateLimit(`inbound:token:${token}`, 60, 3600)
    if (!inboundLimit.allowed) return tooManyRequests(inboundLimit.retryAfter)

    const orgs = (await sql`
      SELECT * FROM organizations WHERE inbox_token = ${token} LIMIT 1
    `) as unknown as Organization[]

    if (orgs.length === 0) {
      return NextResponse.json({ error: 'Unknown inbox' }, { status: 404 })
    }
    const org = orgs[0]

    const raw = (await request.json()) as Record<string, unknown>
    const payload = normalizePayload(raw)

    const from = (payload.from ?? '').slice(0, 320)
    const subject = (payload.subject ?? '').slice(0, 300)
    const bodyText = (payload.text ?? payload.body ?? (payload.html ? stripHtml(payload.html) : '')).slice(0, 10000)

    if (!from || (!subject && !bodyText)) {
      return NextResponse.json({ error: 'Missing from/subject/body' }, { status: 400 })
    }

    const senderEmail = parseFromAddress(from)
    const extracted = await extractTicketFromEmail(from, subject, bodyText)

    let propertyId: string | null = null
    let unitId: string | null = null
    if (extracted.property_hint) {
      const hint = `%${extracted.property_hint.toLowerCase()}%`
      const props = (await sql`
        SELECT * FROM properties
        WHERE organization_id = ${org.id}
          AND (LOWER(name) LIKE ${hint} OR LOWER(address) LIKE ${hint})
        LIMIT 1
      `) as unknown as Property[]
      if (props.length > 0) {
        propertyId = props[0].id
        if (extracted.unit_hint) {
          const unitMatch = (await sql`
            SELECT id FROM units
            WHERE property_id = ${propertyId}
              AND LOWER(unit_number) = LOWER(${extracted.unit_hint})
            LIMIT 1
          `) as unknown as Array<{ id: string }>
          if (unitMatch.length > 0) unitId = unitMatch[0].id
        }
      }
    }

    const ticket = await createTicketWithAI({
      organization_id: org.id,
      title: extracted.title,
      description: extracted.description,
      property_id: propertyId,
      unit_id: unitId,
      tenant_name: extracted.tenant_name,
      tenant_email: extracted.tenant_email ?? senderEmail,
      tenant_phone: extracted.tenant_phone,
      source: 'email',
    })

    return NextResponse.json({ success: true, ticket_id: ticket.id }, { status: 201 })
  } catch (error) {
    console.error('Inbound email error:', error)
    return NextResponse.json({ error: 'Internal server error' }, { status: 500 })
  }
}

export async function GET() {
  return NextResponse.json({
    info: 'POST a JSON email payload to this URL to create a maintenance ticket.',
    expected_fields: ['from', 'subject', 'text or html'],
  })
}

import { NextRequest, NextResponse } from 'next/server'
import { sql } from '@/lib/db'
import { logActivity } from '@/lib/auth'
import { sendSlaBreachEmail, type SlaBreachItem } from '@/lib/email'

type BreachRow = {
  id: string
  organization_id: string
  title: string
  urgency: string
  sla_due_at: string
  property_name: string | null
  vendor_name: string | null
}

type ManagerRow = { organization_id: string; email: string }

export async function GET(request: NextRequest) {
  // Vercel Cron sends `Authorization: Bearer <CRON_SECRET>` when CRON_SECRET is set.
  const cronSecret = process.env.CRON_SECRET
  if (cronSecret) {
    const auth = request.headers.get('authorization')
    if (auth !== `Bearer ${cronSecret}`) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
    }
  }

  try {
    // Tickets past their SLA, not completed/cancelled, and not yet alerted.
    const breaches = (await sql`
      SELECT
        t.id,
        t.organization_id,
        t.title,
        t.urgency,
        t.sla_due_at,
        p.name AS property_name,
        v.name AS vendor_name
      FROM maintenance_tickets t
      LEFT JOIN properties p ON p.id = t.property_id
      LEFT JOIN vendors v ON v.id = t.assigned_vendor_id
      WHERE t.sla_due_at IS NOT NULL
        AND t.sla_due_at < NOW()
        AND t.status NOT IN ('completed', 'cancelled')
        AND NOT EXISTS (
          SELECT 1 FROM activity_logs a
          WHERE a.ticket_id = t.id AND a.action_type = 'sla_breach'
        )
      ORDER BY t.organization_id, t.sla_due_at ASC
    `) as unknown as BreachRow[]

    if (breaches.length === 0) {
      return NextResponse.json({ checked: true, breaches: 0, notified: 0 })
    }

    // Group breaches by organization
    const byOrg = new Map<string, BreachRow[]>()
    for (const b of breaches) {
      const list = byOrg.get(b.organization_id) ?? []
      list.push(b)
      byOrg.set(b.organization_id, list)
    }

    const orgIds = [...byOrg.keys()]
    const managers = (await sql`
      SELECT organization_id, email FROM users
      WHERE organization_id = ANY(${orgIds})
        AND role IN ('admin', 'manager')
    `) as unknown as ManagerRow[]

    const managersByOrg = new Map<string, string[]>()
    for (const m of managers) {
      const list = managersByOrg.get(m.organization_id) ?? []
      list.push(m.email)
      managersByOrg.set(m.organization_id, list)
    }

    const appUrl = process.env.NEXT_PUBLIC_APP_URL || `${request.nextUrl.protocol}//${request.nextUrl.host}`
    const now = Date.now()
    let notified = 0

    for (const [orgId, orgBreaches] of byOrg) {
      const items: SlaBreachItem[] = orgBreaches.map(b => ({
        id: b.id,
        title: b.title,
        urgency: b.urgency,
        hoursOverdue: Math.max(1, Math.round((now - new Date(b.sla_due_at).getTime()) / (60 * 60 * 1000))),
        propertyName: b.property_name,
        vendorName: b.vendor_name,
      }))

      const emails = managersByOrg.get(orgId) ?? []
      for (const email of emails) {
        const ok = await sendSlaBreachEmail(email, items, appUrl)
        if (ok) notified++
      }

      // Mark each ticket as alerted so we don't re-notify on the next run
      for (const b of orgBreaches) {
        await logActivity({
          organizationId: orgId,
          ticketId: b.id,
          actionType: 'sla_breach',
          description: `SLA breached for "${b.title}" (${b.urgency})`,
          metadata: { urgency: b.urgency, sla_due_at: b.sla_due_at },
        })
      }
    }

    return NextResponse.json({ checked: true, breaches: breaches.length, notified })
  } catch (error) {
    console.error('SLA check error:', error)
    return NextResponse.json({ error: 'SLA check failed' }, { status: 500 })
  }
}

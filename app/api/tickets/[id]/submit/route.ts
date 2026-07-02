import { NextRequest, NextResponse } from 'next/server'
import { getSession } from '@/lib/auth'
import { sql, type Organization, type MaintenanceTicket } from '@/lib/db'
import { triageAndDispatch } from '@/lib/tickets'
import { checkResourceLimit, getEffectivePlan, getUsage } from '@/lib/plans'

// Promotes a saved draft into a live ticket: clears the draft flag, then runs
// AI triage + vendor auto-dispatch (and the notifications that flow from it).
export async function POST(
  _request: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const session = await getSession()
    if (!session?.user.organization_id) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
    }
    const orgId = session.user.organization_id
    const { id } = await params

    const rows = (await sql`
      SELECT * FROM maintenance_tickets WHERE id = ${id} AND organization_id = ${orgId}
    `) as unknown as MaintenanceTicket[]
    if (rows.length === 0) {
      return NextResponse.json({ error: 'Ticket not found' }, { status: 404 })
    }
    const ticket = rows[0]
    if (!ticket.is_draft) {
      return NextResponse.json({ error: 'Ticket is already submitted' }, { status: 400 })
    }

    // Now that it becomes a real ticket, enforce the monthly limit.
    const orgRows = (await sql`
      SELECT * FROM organizations WHERE id = ${orgId}
    `) as unknown as Organization[]
    const usage = await getUsage(orgId)
    const check = checkResourceLimit(getEffectivePlan(orgRows[0]), 'tickets_per_month', usage.tickets_this_month)
    if (!check.allowed) {
      return NextResponse.json(
        {
          error: `You've hit your monthly ticket limit (${check.current}/${check.limit}). Upgrade to ${check.upgrade_to ?? 'a higher plan'} to submit this draft.`,
          limit_exceeded: true,
          upgrade_to: check.upgrade_to,
        },
        { status: 402 }
      )
    }

    await sql`UPDATE maintenance_tickets SET is_draft = FALSE, updated_at = NOW() WHERE id = ${id}`
    ticket.is_draft = false

    await triageAndDispatch(ticket)

    return NextResponse.json({ submitted: true, ticket_id: id })
  } catch (error) {
    console.error('Error submitting draft:', error)
    return NextResponse.json({ error: 'Internal server error' }, { status: 500 })
  }
}

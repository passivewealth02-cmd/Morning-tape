import 'server-only'
import { sql, type MaintenanceTicket, type Vendor, type TicketUrgency } from './db'
import { analyzeMaintenanceTicket, recommendVendors } from './anthropic'
import { sendVendorAssignmentEmail, sendTenantVendorAssignedEmail } from './email'
import { logActivity } from './auth'

// Target first-response/resolution window per urgency, in hours
const SLA_HOURS: Record<TicketUrgency, number> = {
  emergency: 4,
  high: 24,
  medium: 72,
  low: 168,
}

export function slaDueDate(urgency: TicketUrgency, from: number = Date.now()): string {
  return new Date(from + SLA_HOURS[urgency] * 60 * 60 * 1000).toISOString()
}

export type CreateTicketInput = {
  organization_id: string
  title: string
  description: string
  urgency?: 'low' | 'medium' | 'high' | 'emergency'
  property_id?: string | null
  unit_id?: string | null
  tenant_name?: string | null
  tenant_email?: string | null
  tenant_phone?: string | null
  created_by?: string | null
  source?: 'manual' | 'email' | 'sms' | 'web'
  draft?: boolean
}

export async function createTicketWithAI(input: CreateTicketInput): Promise<MaintenanceTicket> {
  const {
    organization_id,
    title,
    description,
    urgency = 'medium',
    property_id = null,
    unit_id = null,
    tenant_name = null,
    tenant_email = null,
    tenant_phone = null,
    created_by = null,
    source = 'manual',
    draft = false,
  } = input

  const inserted = (await sql`
    INSERT INTO maintenance_tickets (
      organization_id, property_id, unit_id, title, description, urgency,
      status, is_draft, tenant_name, tenant_email, tenant_phone, created_by, sla_due_at
    )
    VALUES (
      ${organization_id}, ${property_id}, ${unit_id}, ${title}, ${description}, ${urgency},
      'new', ${draft}, ${tenant_name}, ${tenant_email}, ${tenant_phone}, ${created_by}, ${slaDueDate(urgency)}
    )
    RETURNING *
  `) as unknown as MaintenanceTicket[]

  const ticket = inserted[0]

  await logActivity({
    organizationId: organization_id,
    ticketId: ticket.id,
    userId: created_by ?? undefined,
    actionType: draft ? 'draft_created' : source === 'email' ? 'ticket_created_email' : 'ticket_created',
    description: draft
      ? `Draft saved: ${title}`
      : `Ticket created${source !== 'manual' ? ` via ${source}` : ''}: ${title}`,
    metadata: { urgency: ticket.urgency, source, draft },
  })

  // Drafts are held for review — no AI triage, dispatch, or notifications until submitted.
  if (draft) return ticket

  return triageAndDispatch(ticket)
}

// Runs AI classification and auto-dispatch on an existing ticket. Used on
// creation and when a saved draft is submitted.
export async function triageAndDispatch(ticket: MaintenanceTicket): Promise<MaintenanceTicket> {
  const { title, description } = ticket
  let analysisOk = false
  try {
    const analysis = await analyzeMaintenanceTicket(title, description)
    // Re-derive the SLA window from the AI-determined urgency, anchored to creation time
    const slaDue = slaDueDate(analysis.urgency, new Date(ticket.created_at).getTime())
    await sql`
      UPDATE maintenance_tickets
      SET
        ai_category = ${analysis.category},
        ai_vendor_type = ${analysis.vendor_type},
        ai_summary = ${analysis.summary},
        ai_escalation_risk = ${analysis.escalation_risk},
        urgency = ${analysis.urgency},
        sla_due_at = ${slaDue}
      WHERE id = ${ticket.id}
    `
    ticket.ai_category = analysis.category
    ticket.ai_vendor_type = analysis.vendor_type
    ticket.ai_summary = analysis.summary
    ticket.ai_escalation_risk = analysis.escalation_risk
    ticket.urgency = analysis.urgency
    ticket.sla_due_at = slaDue
    analysisOk = true
  } catch (aiError) {
    console.error('AI analysis failed:', aiError)
  }

  if (analysisOk && ticket.ai_vendor_type) {
    try {
      await autoAssignVendor(ticket)
    } catch (assignError) {
      console.error('Auto-assign failed:', assignError)
    }
  }

  return ticket
}

async function autoAssignVendor(ticket: MaintenanceTicket): Promise<void> {
  const vendorType = ticket.ai_vendor_type
  const category = ticket.ai_category
  if (!vendorType && !category) return

  const typeLike = vendorType ? `%${vendorType}%` : null
  const categoryLike = category ? `%${category}%` : null

  let vendors = (await sql`
    SELECT * FROM vendors
    WHERE organization_id = ${ticket.organization_id}
      AND availability != 'unavailable'
      AND (
        (${typeLike}::text IS NOT NULL AND LOWER(trade_type) LIKE LOWER(${typeLike}))
        OR (${categoryLike}::text IS NOT NULL AND LOWER(trade_type) LIKE LOWER(${categoryLike}))
      )
    ORDER BY
      CASE availability WHEN 'available' THEN 0 WHEN 'busy' THEN 1 ELSE 2 END,
      rating DESC
    LIMIT 10
  `) as unknown as Vendor[]

  // Fall back to any available vendor in the org if no trade match
  if (vendors.length === 0) {
    vendors = (await sql`
      SELECT * FROM vendors
      WHERE organization_id = ${ticket.organization_id}
        AND availability != 'unavailable'
      ORDER BY
        CASE availability WHEN 'available' THEN 0 WHEN 'busy' THEN 1 ELSE 2 END,
        rating DESC
      LIMIT 10
    `) as unknown as Vendor[]
  }

  if (vendors.length === 0) return

  const ranked = await recommendVendors(
    {
      category: ticket.ai_category ?? 'general',
      urgency: ticket.urgency,
      vendor_type: vendorType,
      summary: ticket.ai_summary ?? ticket.title,
      escalation_risk: ticket.ai_escalation_risk,
      escalation_reason: null,
    },
    vendors.map(v => ({
      id: v.id,
      name: v.name,
      trade_type: v.trade_type,
      rating: v.rating,
      availability: v.availability,
    }))
  )

  const topId = ranked[0] ?? vendors[0].id
  const top = vendors.find(v => v.id === topId) ?? vendors[0]

  await sql`
    UPDATE maintenance_tickets
    SET assigned_vendor_id = ${top.id}, status = 'assigned', updated_at = NOW()
    WHERE id = ${ticket.id}
  `
  ticket.assigned_vendor_id = top.id
  ticket.status = 'assigned'

  await logActivity({
    organizationId: ticket.organization_id,
    ticketId: ticket.id,
    actionType: 'vendor_auto_assigned',
    description: `Auto-assigned to ${top.name}`,
    metadata: { vendor_id: top.id, vendor_name: top.name },
  })

  if (top.email) {
    await sendVendorAssignmentEmail(top.email, top.name, ticket.title, ticket.id)
  }

  if (ticket.tenant_email) {
    await sendTenantVendorAssignedEmail(
      ticket.tenant_email,
      ticket.tenant_name,
      ticket.title,
      top.name,
    )
  }
}

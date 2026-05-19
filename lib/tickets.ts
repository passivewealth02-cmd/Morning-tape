import 'server-only'
import { sql, type MaintenanceTicket, type Vendor } from './db'
import { analyzeMaintenanceTicket, recommendVendors } from './anthropic'
import { sendVendorAssignmentEmail } from './email'
import { logActivity } from './auth'

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
  source?: 'manual' | 'email' | 'sms'
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
  } = input

  const inserted = (await sql`
    INSERT INTO maintenance_tickets (
      organization_id, property_id, unit_id, title, description, urgency,
      status, tenant_name, tenant_email, tenant_phone, created_by
    )
    VALUES (
      ${organization_id}, ${property_id}, ${unit_id}, ${title}, ${description}, ${urgency},
      'new', ${tenant_name}, ${tenant_email}, ${tenant_phone}, ${created_by}
    )
    RETURNING *
  `) as unknown as MaintenanceTicket[]

  const ticket = inserted[0]

  await logActivity({
    organizationId: organization_id,
    ticketId: ticket.id,
    userId: created_by ?? undefined,
    actionType: source === 'email' ? 'ticket_created_email' : 'ticket_created',
    description: `Ticket created${source !== 'manual' ? ` via ${source}` : ''}: ${title}`,
    metadata: { urgency: ticket.urgency, source },
  })

  let analysisOk = false
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
    await sendVendorAssignmentEmail(
      top.email,
      top.name,
      ticket.title,
      ticket.id
    )
  }
}

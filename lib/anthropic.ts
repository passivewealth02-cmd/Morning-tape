import 'server-only'
import Anthropic from '@anthropic-ai/sdk'

const anthropic = new Anthropic({
  apiKey: process.env.ANTHROPIC_API_KEY,
})

export type TicketAnalysis = {
  category: string
  urgency: 'low' | 'medium' | 'high' | 'emergency'
  vendor_type: string
  summary: string
  escalation_risk: boolean
  escalation_reason: string | null
}

export async function analyzeMaintenanceTicket(
  title: string,
  description: string
): Promise<TicketAnalysis> {
  const prompt = `You are an AI assistant for a property maintenance coordination platform. Analyze the following maintenance request and return structured data.

Title: ${title}
Description: ${description}

Return a JSON object with exactly these fields:
{
  "category": "one of: plumbing, electrical, hvac, appliance, structural, pest_control, landscaping, cleaning, security, general",
  "urgency": "one of: low, medium, high, emergency",
  "vendor_type": "specific trade needed, e.g. 'plumber', 'electrician', 'hvac_technician', 'handyman', etc.",
  "summary": "concise 1-2 sentence operational summary for the property manager",
  "escalation_risk": true or false,
  "escalation_reason": "string if escalation_risk is true, null otherwise"
}

Urgency guidelines:
- emergency: immediate safety risk, flooding, no heat in winter, gas leak
- high: significant impact on livability, needs same-day response
- medium: important but not urgent, should be resolved within 48-72 hours
- low: minor issue, can be scheduled within a week

Escalation risk is true if: the request contains angry language, mentions repeated issues, threatens legal action, or describes a health/safety hazard.

Return ONLY valid JSON, no markdown.`

  const message = await anthropic.messages.create({
    model: 'claude-haiku-4-5-20251001',
    max_tokens: 512,
    messages: [{ role: 'user', content: prompt }],
  })

  const content = message.content[0]
  if (content.type !== 'text') throw new Error('Unexpected response type')

  try {
    return JSON.parse(content.text) as TicketAnalysis
  } catch {
    // Fallback if parsing fails
    return {
      category: 'general',
      urgency: 'medium',
      vendor_type: 'handyman',
      summary: `${title}: ${description.slice(0, 100)}`,
      escalation_risk: false,
      escalation_reason: null,
    }
  }
}

export async function recommendVendors(
  ticketAnalysis: TicketAnalysis,
  vendors: Array<{ id: string; name: string; trade_type: string; rating: number; availability: string }>
): Promise<string[]> {
  if (vendors.length === 0) return []

  const vendorList = vendors
    .map(v => `ID: ${v.id}, Name: ${v.name}, Trade: ${v.trade_type}, Rating: ${v.rating}, Availability: ${v.availability}`)
    .join('\n')

  const prompt = `A maintenance ticket needs a vendor. Here are the details:
Category: ${ticketAnalysis.category}
Required vendor type: ${ticketAnalysis.vendor_type}
Urgency: ${ticketAnalysis.urgency}

Available vendors:
${vendorList}

Return a JSON array of vendor IDs ordered by best match (best first, max 3). Consider trade type match, availability, and rating.
Return ONLY a JSON array like: ["id1", "id2"] — no markdown.`

  const message = await anthropic.messages.create({
    model: 'claude-haiku-4-5-20251001',
    max_tokens: 256,
    messages: [{ role: 'user', content: prompt }],
  })

  const content = message.content[0]
  if (content.type !== 'text') return []

  try {
    return JSON.parse(content.text) as string[]
  } catch {
    return vendors.slice(0, 3).map(v => v.id)
  }
}

export type InboundEmailExtraction = {
  title: string
  description: string
  tenant_name: string | null
  tenant_email: string | null
  tenant_phone: string | null
  property_hint: string | null
  unit_hint: string | null
}

export async function extractTicketFromEmail(
  from: string,
  subject: string,
  body: string
): Promise<InboundEmailExtraction> {
  const prompt = `You are processing an inbound email reporting a property maintenance issue. Extract structured ticket data.

From: ${from}
Subject: ${subject}
Body:
${body.slice(0, 4000)}

Return a JSON object with exactly these fields:
{
  "title": "short ticket title, max 80 chars, summarizing the issue",
  "description": "clean issue description with email signatures/quoted replies/disclaimers stripped",
  "tenant_name": "the sender's full name if discoverable, else null",
  "tenant_email": "the sender's email address",
  "tenant_phone": "phone number if mentioned in body or signature, else null",
  "property_hint": "property name or street address if mentioned, else null",
  "unit_hint": "unit/apartment number if mentioned (e.g. 'Apt 4B', '#203'), else null"
}

Return ONLY valid JSON, no markdown.`

  const message = await anthropic.messages.create({
    model: 'claude-haiku-4-5-20251001',
    max_tokens: 1024,
    messages: [{ role: 'user', content: prompt }],
  })

  const content = message.content[0]
  if (content.type !== 'text') throw new Error('Unexpected response type')

  try {
    return JSON.parse(content.text) as InboundEmailExtraction
  } catch {
    return {
      title: subject.slice(0, 80) || 'Maintenance request',
      description: body.slice(0, 1000),
      tenant_name: null,
      tenant_email: from,
      tenant_phone: null,
      property_hint: null,
      unit_hint: null,
    }
  }
}

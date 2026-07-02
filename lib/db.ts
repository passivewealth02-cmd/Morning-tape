import { neon } from '@neondatabase/serverless'
import { Pool } from 'pg'

const dbUrl = process.env.DATABASE_URL!
const isLocal = dbUrl.includes('localhost') || dbUrl.includes('127.0.0.1')

let sql: ReturnType<typeof neon>

if (isLocal) {
  const pool = new Pool({ connectionString: dbUrl })
  // Match the neon tagged-template API
  sql = ((strings: TemplateStringsArray, ...values: unknown[]) => {
    let text = ''
    strings.forEach((s, i) => {
      text += s
      if (i < values.length) text += `$${i + 1}`
    })
    return pool.query(text, values as unknown[]).then(r => r.rows) as ReturnType<ReturnType<typeof neon>>
  }) as ReturnType<typeof neon>
} else {
  sql = neon(dbUrl)
}

export { sql }

export type OrganizationPlan = 'trial' | 'starter' | 'growth' | 'pro'
export type PlanStatus = 'active' | 'past_due' | 'canceled' | 'trialing' | 'incomplete'

export type Organization = {
  id: string
  name: string
  slug: string
  inbox_token: string | null
  plan: OrganizationPlan
  plan_status: PlanStatus
  trial_ends_at: string | null
  stripe_customer_id: string | null
  stripe_subscription_id: string | null
  stripe_price_id: string | null
  created_at: string
  updated_at: string
}

export type User = {
  id: string
  organization_id: string | null
  email: string
  name: string | null
  role: 'admin' | 'manager' | 'coordinator' | 'vendor'
  created_at: string
  updated_at: string
}

export type Property = {
  id: string
  organization_id: string
  name: string
  address: string
  city: string | null
  province: string | null
  unit_count: number
  created_at: string
  updated_at: string
}

export type Unit = {
  id: string
  property_id: string
  unit_number: string
  tenant_name: string | null
  tenant_email: string | null
  tenant_phone: string | null
  created_at: string
}

export type Vendor = {
  id: string
  organization_id: string
  name: string
  trade_type: string
  email: string | null
  phone: string | null
  service_areas: string[]
  availability: 'available' | 'busy' | 'unavailable'
  rating: number
  insurance_status: 'verified' | 'expired' | 'unknown'
  notes: string | null
  created_at: string
  updated_at: string
}

export type TicketStatus = 'new' | 'assigned' | 'in_progress' | 'waiting' | 'completed' | 'cancelled'
export type TicketUrgency = 'low' | 'medium' | 'high' | 'emergency'

export type MaintenanceTicket = {
  id: string
  organization_id: string
  property_id: string | null
  unit_id: string | null
  title: string
  description: string
  status: TicketStatus
  urgency: TicketUrgency
  is_draft: boolean
  ai_category: string | null
  ai_vendor_type: string | null
  ai_summary: string | null
  ai_escalation_risk: boolean
  assigned_vendor_id: string | null
  created_by: string | null
  tenant_name: string | null
  tenant_email: string | null
  tenant_phone: string | null
  first_response_at: string | null
  assigned_at: string | null
  completed_at: string | null
  sla_due_at: string | null
  created_at: string
  updated_at: string
  // Joined fields
  property_name?: string
  property_address?: string
  unit_number?: string
  vendor_name?: string
  vendor_trade_type?: string
}

export type TicketMessage = {
  id: string
  ticket_id: string
  sender_type: 'manager' | 'vendor' | 'tenant' | 'system'
  sender_id: string | null
  message: string
  is_internal: boolean
  created_at: string
  sender_name?: string
}

export type TicketFile = {
  id: string
  ticket_id: string
  file_url: string
  file_name: string | null
  file_type: string | null
  uploaded_by: string | null
  created_at: string
}

export type ActivityLog = {
  id: string
  organization_id: string
  ticket_id: string | null
  user_id: string | null
  action_type: string
  description: string
  metadata: Record<string, unknown> | null
  created_at: string
  user_name?: string
}

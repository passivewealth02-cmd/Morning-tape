import 'server-only'
import { sql, type Organization, type OrganizationPlan } from './db'

export const TRIAL_DAYS = 14
const OVERAGE_PERCENT = 10 // allow 10% over before hard-blocking
const WARN_PERCENT = 80 // surface a warning at 80%

type Limits = {
  tickets_per_month: number // -1 = unlimited
  properties: number
  vendors: number
  ai_dispatch: boolean
  sla_tracking: boolean
  white_label: boolean
  api_access: boolean
}

export const PLAN_LIMITS: Record<OrganizationPlan, Limits> = {
  trial: {
    tickets_per_month: -1,
    properties: 10,
    vendors: -1,
    ai_dispatch: true,
    sla_tracking: true,
    white_label: false,
    api_access: false,
  },
  starter: {
    tickets_per_month: 100,
    properties: 1,
    vendors: 5,
    ai_dispatch: false,
    sla_tracking: false,
    white_label: false,
    api_access: false,
  },
  growth: {
    tickets_per_month: -1,
    properties: 10,
    vendors: -1,
    ai_dispatch: true,
    sla_tracking: true,
    white_label: false,
    api_access: false,
  },
  pro: {
    tickets_per_month: -1,
    properties: -1,
    vendors: -1,
    ai_dispatch: true,
    sla_tracking: true,
    white_label: true,
    api_access: true,
  },
}

export const PLAN_LABELS: Record<OrganizationPlan, string> = {
  trial: 'Trial',
  starter: 'Starter',
  growth: 'Growth',
  pro: 'Pro',
}

export function getEffectivePlan(org: Pick<Organization, 'plan' | 'plan_status' | 'trial_ends_at'>): OrganizationPlan {
  // Legacy DB-only trial: if expired with no paid plan, fall back to starter
  if (org.plan === 'trial' && org.trial_ends_at) {
    const expired = new Date(org.trial_ends_at).getTime() < Date.now()
    if (expired) return 'starter'
  }
  return org.plan
}

// True if the org is currently inside an active trial window (Stripe-managed or legacy DB)
export function isTrialActive(org: Pick<Organization, 'plan' | 'plan_status' | 'trial_ends_at'>): boolean {
  if (!org.trial_ends_at) return false
  const notExpired = new Date(org.trial_ends_at).getTime() > Date.now()
  // Stripe-managed trial: plan is already the chosen plan, status is 'trialing'
  if (org.plan_status === 'trialing') return notExpired
  // Legacy DB-only trial
  if (org.plan === 'trial') return notExpired
  return false
}

export function trialDaysLeft(org: Pick<Organization, 'plan' | 'plan_status' | 'trial_ends_at'>): number {
  if (!isTrialActive(org)) return 0
  const ms = new Date(org.trial_ends_at!).getTime() - Date.now()
  return Math.max(0, Math.ceil(ms / (1000 * 60 * 60 * 24)))
}

export type UsageSnapshot = {
  tickets_this_month: number
  properties: number
  vendors: number
}

export async function getUsage(organizationId: string): Promise<UsageSnapshot> {
  const monthStart = new Date()
  monthStart.setUTCDate(1)
  monthStart.setUTCHours(0, 0, 0, 0)

  const [tickets, properties, vendors] = (await Promise.all([
    sql`SELECT COUNT(*)::int AS c FROM maintenance_tickets WHERE organization_id = ${organizationId} AND created_at >= ${monthStart.toISOString()}`,
    sql`SELECT COUNT(*)::int AS c FROM properties WHERE organization_id = ${organizationId}`,
    sql`SELECT COUNT(*)::int AS c FROM vendors WHERE organization_id = ${organizationId}`,
  ])) as unknown as Array<Array<{ c: number }>>

  return {
    tickets_this_month: tickets[0]?.c ?? 0,
    properties: properties[0]?.c ?? 0,
    vendors: vendors[0]?.c ?? 0,
  }
}

export type LimitCheck = {
  allowed: boolean
  warn: boolean
  limit: number
  current: number
  percentage: number
  upgrade_to?: OrganizationPlan
}

export function checkResourceLimit(
  plan: OrganizationPlan,
  resource: 'tickets_per_month' | 'properties' | 'vendors',
  current: number
): LimitCheck {
  const limit = PLAN_LIMITS[plan][resource]
  if (limit === -1) {
    return { allowed: true, warn: false, limit: -1, current, percentage: 0 }
  }

  const percentage = (current / limit) * 100
  const allowed = percentage < 100 + OVERAGE_PERCENT
  const warn = percentage >= WARN_PERCENT

  let upgrade_to: OrganizationPlan | undefined
  if (!allowed || warn) {
    if (plan === 'trial' || plan === 'starter') upgrade_to = 'growth'
    else if (plan === 'growth') upgrade_to = 'pro'
  }

  return { allowed, warn, limit, current, percentage, upgrade_to }
}

import 'server-only'
import Stripe from 'stripe'
import type { OrganizationPlan } from './db'

export const stripe = process.env.STRIPE_SECRET_KEY
  ? new Stripe(process.env.STRIPE_SECRET_KEY, {
      apiVersion: '2026-04-22.dahlia',
      typescript: true,
    })
  : (null as unknown as Stripe)

export const PLAN_PRICE_IDS: Record<Exclude<OrganizationPlan, 'trial'>, string | undefined> = {
  starter: process.env.STRIPE_PRICE_STARTER,
  growth: process.env.STRIPE_PRICE_GROWTH,
  pro: process.env.STRIPE_PRICE_PRO,
}

export function planForPriceId(priceId: string | null | undefined): OrganizationPlan | null {
  if (!priceId) return null
  for (const [plan, id] of Object.entries(PLAN_PRICE_IDS)) {
    if (id === priceId) return plan as OrganizationPlan
  }
  return null
}

export function isStripeConfigured(): boolean {
  return Boolean(process.env.STRIPE_SECRET_KEY)
}

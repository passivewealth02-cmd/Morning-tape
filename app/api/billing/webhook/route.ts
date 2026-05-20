import { NextRequest, NextResponse } from 'next/server'
import Stripe from 'stripe'
import { sql } from '@/lib/db'
import { stripe, planForPriceId, isStripeConfigured } from '@/lib/stripe'

export async function POST(request: NextRequest) {
  if (!isStripeConfigured() || !process.env.STRIPE_WEBHOOK_SECRET) {
    return NextResponse.json({ error: 'Webhook not configured' }, { status: 503 })
  }

  const body = await request.text()
  const signature = request.headers.get('stripe-signature')
  if (!signature) {
    return NextResponse.json({ error: 'Missing signature' }, { status: 400 })
  }

  let event: Stripe.Event
  try {
    event = stripe.webhooks.constructEvent(body, signature, process.env.STRIPE_WEBHOOK_SECRET)
  } catch (err) {
    console.error('Invalid webhook signature:', err)
    return NextResponse.json({ error: 'Invalid signature' }, { status: 400 })
  }

  try {
    switch (event.type) {
      case 'checkout.session.completed': {
        const session = event.data.object as Stripe.Checkout.Session
        const orgId = session.metadata?.organization_id
        if (!orgId || !session.subscription) break

        const subscriptionId = typeof session.subscription === 'string'
          ? session.subscription
          : session.subscription.id

        const subscription = await stripe.subscriptions.retrieve(subscriptionId)
        const priceId = subscription.items.data[0]?.price.id
        const plan = planForPriceId(priceId)

        if (plan) {
          // Preserve Stripe's trial_end so the dashboard can show the countdown
          const trialEndsAt = subscription.trial_end
            ? new Date(subscription.trial_end * 1000).toISOString()
            : null

          await sql`
            UPDATE organizations
            SET
              plan = ${plan},
              plan_status = ${subscription.status},
              stripe_subscription_id = ${subscription.id},
              stripe_price_id = ${priceId},
              trial_ends_at = ${trialEndsAt},
              updated_at = NOW()
            WHERE id = ${orgId}
          `
        }
        break
      }

      case 'customer.subscription.updated':
      case 'customer.subscription.created': {
        const subscription = event.data.object as Stripe.Subscription
        const orgId = subscription.metadata?.organization_id
        if (!orgId) break

        const priceId = subscription.items.data[0]?.price.id
        const plan = planForPriceId(priceId)
        if (plan) {
          await sql`
            UPDATE organizations
            SET
              plan = ${plan},
              plan_status = ${subscription.status},
              stripe_subscription_id = ${subscription.id},
              stripe_price_id = ${priceId},
              updated_at = NOW()
            WHERE id = ${orgId}
          `
        }
        break
      }

      case 'customer.subscription.deleted': {
        const subscription = event.data.object as Stripe.Subscription
        const orgId = subscription.metadata?.organization_id
        if (!orgId) break

        await sql`
          UPDATE organizations
          SET
            plan = 'starter',
            plan_status = 'canceled',
            stripe_subscription_id = NULL,
            stripe_price_id = NULL,
            updated_at = NOW()
          WHERE id = ${orgId}
        `
        break
      }

      case 'invoice.payment_failed': {
        const invoice = event.data.object as Stripe.Invoice
        const customerId = typeof invoice.customer === 'string' ? invoice.customer : invoice.customer?.id
        if (!customerId) break
        await sql`
          UPDATE organizations
          SET plan_status = 'past_due', updated_at = NOW()
          WHERE stripe_customer_id = ${customerId}
        `
        break
      }
    }

    return NextResponse.json({ received: true })
  } catch (error) {
    console.error('Webhook handler error:', error)
    return NextResponse.json({ error: 'Webhook handler failed' }, { status: 500 })
  }
}

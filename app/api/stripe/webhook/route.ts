import { NextRequest, NextResponse } from 'next/server'
import { stripe } from '@/lib/stripe'
import { sql } from '@/lib/db'
import Stripe from 'stripe'

export async function POST(request: NextRequest) {
  const body = await request.text()
  const signature = request.headers.get('stripe-signature')

  if (!signature) {
    return NextResponse.json({ error: 'No signature' }, { status: 400 })
  }

  let event: Stripe.Event

  try {
    event = stripe.webhooks.constructEvent(
      body,
      signature,
      process.env.STRIPE_WEBHOOK_SECRET!
    )
  } catch (err) {
    console.error('Webhook signature verification failed:', err)
    return NextResponse.json({ error: 'Invalid signature' }, { status: 400 })
  }

  try {
    switch (event.type) {
      case 'checkout.session.completed': {
        const session = event.data.object as Stripe.Checkout.Session
        
        if (session.mode === 'subscription' && session.subscription) {
          const userId = session.metadata?.userId
          const planType = session.metadata?.planType as 'trader' | 'professional'
          
          if (!userId || !planType) {
            console.error('Missing userId or planType in session metadata')
            break
          }

          const subscription = await stripe.subscriptions.retrieve(
            session.subscription as string
          )

          // Create or update subscription record
          await sql`
            INSERT INTO subscriptions (
              user_id, 
              stripe_customer_id, 
              stripe_subscription_id, 
              plan, 
              status, 
              current_period_end
            )
            VALUES (
              ${userId},
              ${session.customer as string},
              ${subscription.id},
              ${planType},
              'active',
              ${new Date(subscription.current_period_end * 1000).toISOString()}
            )
            ON CONFLICT (user_id) 
            DO UPDATE SET
              stripe_subscription_id = ${subscription.id},
              plan = ${planType},
              status = 'active',
              current_period_end = ${new Date(subscription.current_period_end * 1000).toISOString()},
              updated_at = NOW()
          `
        }
        break
      }

      case 'customer.subscription.updated': {
        const subscription = event.data.object as Stripe.Subscription
        const status = subscription.status

        // Map Stripe status to our status
        let ourStatus: 'active' | 'canceled' | 'past_due' | 'trialing'
        switch (status) {
          case 'active':
            ourStatus = 'active'
            break
          case 'canceled':
            ourStatus = 'canceled'
            break
          case 'past_due':
            ourStatus = 'past_due'
            break
          case 'trialing':
            ourStatus = 'trialing'
            break
          default:
            ourStatus = 'canceled'
        }

        await sql`
          UPDATE subscriptions
          SET 
            status = ${ourStatus},
            current_period_end = ${new Date(subscription.current_period_end * 1000).toISOString()},
            updated_at = NOW()
          WHERE stripe_subscription_id = ${subscription.id}
        `
        break
      }

      case 'customer.subscription.deleted': {
        const subscription = event.data.object as Stripe.Subscription

        await sql`
          UPDATE subscriptions
          SET 
            status = 'canceled',
            updated_at = NOW()
          WHERE stripe_subscription_id = ${subscription.id}
        `
        break
      }

      default:
        console.log(`Unhandled event type: ${event.type}`)
    }

    return NextResponse.json({ received: true })
  } catch (error) {
    console.error('Error processing webhook:', error)
    return NextResponse.json(
      { error: 'Webhook handler failed' },
      { status: 500 }
    )
  }
}

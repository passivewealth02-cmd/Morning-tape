'use server'

import { stripe } from '@/lib/stripe'
import { PRODUCTS, getProductById } from '@/lib/products'
import { getSession } from '@/lib/auth'
import { sql } from '@/lib/db'

export async function startCheckoutSession(productId: string) {
  const product = getProductById(productId)
  if (!product) {
    throw new Error(`Product with id "${productId}" not found`)
  }

  // Get current user session
  const session = await getSession()
  if (!session) {
    throw new Error('You must be logged in to subscribe')
  }

  const { user } = session

  // Check if user already has a Stripe customer ID
  let stripeCustomerId: string | null = null
  const existingSub = await sql`
    SELECT stripe_customer_id FROM subscriptions
    WHERE user_id = ${user.id}
    ORDER BY created_at DESC
    LIMIT 1
  `
  
  if (existingSub.length > 0 && existingSub[0].stripe_customer_id) {
    stripeCustomerId = existingSub[0].stripe_customer_id
  } else {
    // Create a new Stripe customer
    const customer = await stripe.customers.create({
      email: user.email,
      metadata: {
        userId: user.id,
      },
    })
    stripeCustomerId = customer.id
  }

  // Create Checkout Session for subscription
  const checkoutSession = await stripe.checkout.sessions.create({
    ui_mode: 'embedded',
    redirect_on_completion: 'never',
    customer: stripeCustomerId,
    line_items: [
      {
        price_data: {
          currency: 'usd',
          product_data: {
            name: product.name,
            description: product.description,
          },
          unit_amount: product.priceInCents,
          recurring: {
            interval: 'month',
          },
        },
        quantity: 1,
      },
    ],
    mode: 'subscription',
    metadata: {
      userId: user.id,
      planType: product.planType,
    },
  })

  return checkoutSession.client_secret
}

export async function getCheckoutSessionStatus(sessionId: string) {
  const session = await stripe.checkout.sessions.retrieve(sessionId)
  return {
    status: session.status,
    customerEmail: session.customer_details?.email,
  }
}

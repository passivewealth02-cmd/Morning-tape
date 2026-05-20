import { NextRequest, NextResponse } from 'next/server'
import { getSession } from '@/lib/auth'
import { sql, type Organization, type OrganizationPlan } from '@/lib/db'
import { stripe, PLAN_PRICE_IDS, isStripeConfigured } from '@/lib/stripe'

export async function POST(request: NextRequest) {
  try {
    if (!isStripeConfigured()) {
      return NextResponse.json({ error: 'Billing is not configured yet.' }, { status: 503 })
    }

    const session = await getSession()
    if (!session) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
    }

    const { user } = session
    if (!user.organization_id) {
      return NextResponse.json({ error: 'No organization' }, { status: 403 })
    }
    if (user.role !== 'admin') {
      return NextResponse.json({ error: 'Only admins can manage billing' }, { status: 403 })
    }

    const { plan } = (await request.json()) as { plan: OrganizationPlan }
    if (plan === 'trial' || !(plan in PLAN_PRICE_IDS)) {
      return NextResponse.json({ error: 'Invalid plan' }, { status: 400 })
    }

    const priceId = PLAN_PRICE_IDS[plan as Exclude<OrganizationPlan, 'trial'>]
    if (!priceId) {
      return NextResponse.json({ error: `Price for ${plan} is not configured` }, { status: 500 })
    }

    const orgRows = (await sql`
      SELECT * FROM organizations WHERE id = ${user.organization_id}
    `) as unknown as Organization[]
    const org = orgRows[0]

    let customerId = org.stripe_customer_id
    if (!customerId) {
      const customer = await stripe.customers.create({
        email: user.email,
        name: org.name,
        metadata: { organization_id: org.id, user_id: user.id },
      })
      customerId = customer.id
      await sql`
        UPDATE organizations SET stripe_customer_id = ${customerId}, updated_at = NOW()
        WHERE id = ${org.id}
      `
    }

    const origin = `${request.nextUrl.protocol}//${request.nextUrl.host}`
    const isNewSubscriber = !org.stripe_subscription_id

    const checkout = await stripe.checkout.sessions.create({
      mode: 'subscription',
      customer: customerId,
      line_items: [{ price: priceId, quantity: 1 }],
      allow_promotion_codes: true,
      success_url: `${origin}/dashboard/settings?billing=success`,
      cancel_url: `${origin}/dashboard/settings?billing=cancelled`,
      subscription_data: {
        metadata: { organization_id: org.id, plan },
        // New subscribers get a 14-day trial; existing subscribers switching plans do not
        ...(isNewSubscriber ? { trial_period_days: 14 } : {}),
      },
      metadata: { organization_id: org.id, plan },
    })

    return NextResponse.json({ url: checkout.url })
  } catch (error) {
    console.error('Checkout error:', error)
    return NextResponse.json({ error: 'Could not start checkout' }, { status: 500 })
  }
}

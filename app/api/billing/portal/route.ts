import { NextRequest, NextResponse } from 'next/server'
import { getSession } from '@/lib/auth'
import { sql, type Organization } from '@/lib/db'
import { stripe, isStripeConfigured } from '@/lib/stripe'

export async function POST(request: NextRequest) {
  try {
    if (!isStripeConfigured()) {
      return NextResponse.json({ error: 'Billing is not configured yet.' }, { status: 503 })
    }

    const session = await getSession()
    if (!session) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })

    const { user } = session
    if (!user.organization_id) return NextResponse.json({ error: 'No organization' }, { status: 403 })
    if (user.role !== 'admin') {
      return NextResponse.json({ error: 'Only admins can manage billing' }, { status: 403 })
    }

    const orgRows = (await sql`
      SELECT * FROM organizations WHERE id = ${user.organization_id}
    `) as unknown as Organization[]
    const org = orgRows[0]

    if (!org.stripe_customer_id) {
      return NextResponse.json({ error: 'No active subscription to manage' }, { status: 400 })
    }

    const origin = `${request.nextUrl.protocol}//${request.nextUrl.host}`
    const portal = await stripe.billingPortal.sessions.create({
      customer: org.stripe_customer_id,
      return_url: `${origin}/dashboard/settings`,
    })

    return NextResponse.json({ url: portal.url })
  } catch (error) {
    console.error('Portal error:', error)
    return NextResponse.json({ error: 'Could not open billing portal' }, { status: 500 })
  }
}

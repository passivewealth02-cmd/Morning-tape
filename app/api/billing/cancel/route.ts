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

    const { reactivate } = (await request.json().catch(() => ({}))) as { reactivate?: boolean }

    const orgRows = (await sql`
      SELECT * FROM organizations WHERE id = ${user.organization_id}
    `) as unknown as Organization[]
    const org = orgRows[0]

    if (!org.stripe_subscription_id) {
      return NextResponse.json({ error: 'No active subscription to cancel' }, { status: 400 })
    }

    // Schedule cancellation at period end (keeps access through the trial/paid period),
    // or undo a previously scheduled cancellation.
    const subscription = await stripe.subscriptions.update(org.stripe_subscription_id, {
      cancel_at_period_end: !reactivate,
    })

    return NextResponse.json({
      success: true,
      cancel_at_period_end: subscription.cancel_at_period_end,
    })
  } catch (error) {
    console.error('Cancel error:', error)
    return NextResponse.json({ error: 'Could not update subscription' }, { status: 500 })
  }
}

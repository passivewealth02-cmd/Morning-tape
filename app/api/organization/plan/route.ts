import { NextRequest, NextResponse } from 'next/server'
import { getSession } from '@/lib/auth'
import { sql } from '@/lib/db'

const ALLOWED_PLANS = ['trial', 'starter', 'growth', 'pro'] as const

export async function POST(request: NextRequest) {
  try {
    const session = await getSession()
    if (!session) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
    }

    const { user } = session
    if (!user.organization_id) {
      return NextResponse.json({ error: 'No organization' }, { status: 403 })
    }
    if (user.role !== 'admin') {
      return NextResponse.json({ error: 'Only admins can change plan' }, { status: 403 })
    }

    const { plan } = await request.json()
    if (!ALLOWED_PLANS.includes(plan)) {
      return NextResponse.json({ error: 'Invalid plan' }, { status: 400 })
    }

    await sql`
      UPDATE organizations
      SET plan = ${plan}, updated_at = NOW()
      WHERE id = ${user.organization_id}
    `

    return NextResponse.json({ success: true, plan })
  } catch (error) {
    console.error('Error updating plan:', error)
    return NextResponse.json({ error: 'Internal server error' }, { status: 500 })
  }
}

import { NextRequest, NextResponse } from 'next/server'
import { getSession } from '@/lib/auth'
import { sql } from '@/lib/db'
import { analyzeMaintenanceTicket, recommendVendors, type TicketAnalysis } from '@/lib/anthropic'

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

    const body = await request.json()

    if (body.title && body.description) {
      const analysis = await analyzeMaintenanceTicket(body.title, body.description)
      return NextResponse.json(analysis)
    }

    if (body.ticket_id && body.vendor_analysis) {
      const vendorAnalysis = body.vendor_analysis as TicketAnalysis

      const vendors = await sql`
        SELECT id, name, trade_type, rating, availability
        FROM vendors
        WHERE organization_id = ${user.organization_id}
          AND availability != 'unavailable'
      `

      const recommendedIds = await recommendVendors(
        vendorAnalysis,
        vendors as Array<{ id: string; name: string; trade_type: string; rating: number; availability: string }>
      )

      return NextResponse.json({ recommended_vendor_ids: recommendedIds })
    }

    return NextResponse.json(
      { error: 'Provide either {title, description} or {ticket_id, vendor_analysis}' },
      { status: 400 }
    )
  } catch (error) {
    console.error('Error in AI categorize:', error)
    return NextResponse.json({ error: 'Internal server error' }, { status: 500 })
  }
}

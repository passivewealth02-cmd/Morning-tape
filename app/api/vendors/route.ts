import { NextRequest, NextResponse } from 'next/server'
import { getSession } from '@/lib/auth'
import { sql } from '@/lib/db'

export async function GET(_request: NextRequest) {
  try {
    const session = await getSession()
    if (!session) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
    }

    const { user } = session
    if (!user.organization_id) {
      return NextResponse.json({ error: 'No organization' }, { status: 403 })
    }

    const vendors = await sql`
      SELECT * FROM vendors
      WHERE organization_id = ${user.organization_id}
      ORDER BY name ASC
    `

    return NextResponse.json(vendors)
  } catch (error) {
    console.error('Error fetching vendors:', error)
    return NextResponse.json({ error: 'Internal server error' }, { status: 500 })
  }
}

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
    const {
      name,
      trade_type,
      email = null,
      phone = null,
      notes = null,
      insurance_status = 'unknown',
    } = body

    if (!name || !trade_type) {
      return NextResponse.json({ error: 'Name and trade type are required' }, { status: 400 })
    }

    const inserted = await sql`
      INSERT INTO vendors (organization_id, name, trade_type, email, phone, notes, insurance_status)
      VALUES (${user.organization_id}, ${name}, ${trade_type}, ${email}, ${phone}, ${notes}, ${insurance_status})
      RETURNING *
    `

    return NextResponse.json(inserted[0], { status: 201 })
  } catch (error) {
    console.error('Error creating vendor:', error)
    return NextResponse.json({ error: 'Internal server error' }, { status: 500 })
  }
}

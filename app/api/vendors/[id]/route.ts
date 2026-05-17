import { NextRequest, NextResponse } from 'next/server'
import { getSession } from '@/lib/auth'
import { sql } from '@/lib/db'

export async function GET(
  _request: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const session = await getSession()
    if (!session) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
    }

    const { user } = session
    if (!user.organization_id) {
      return NextResponse.json({ error: 'No organization' }, { status: 403 })
    }

    const { id } = await params

    const vendors = await sql`
      SELECT * FROM vendors
      WHERE id = ${id} AND organization_id = ${user.organization_id}
    `

    if (vendors.length === 0) {
      return NextResponse.json({ error: 'Vendor not found' }, { status: 404 })
    }

    return NextResponse.json(vendors[0])
  } catch (error) {
    console.error('Error fetching vendor:', error)
    return NextResponse.json({ error: 'Internal server error' }, { status: 500 })
  }
}

export async function PATCH(
  request: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const session = await getSession()
    if (!session) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
    }

    const { user } = session
    if (!user.organization_id) {
      return NextResponse.json({ error: 'No organization' }, { status: 403 })
    }

    const { id } = await params

    const existing = await sql`
      SELECT * FROM vendors
      WHERE id = ${id} AND organization_id = ${user.organization_id}
    `

    if (existing.length === 0) {
      return NextResponse.json({ error: 'Vendor not found' }, { status: 404 })
    }

    const current = existing[0]
    const body = await request.json()
    const {
      name,
      trade_type,
      email,
      phone,
      notes,
      insurance_status,
      availability,
      rating,
    } = body

    const updated = await sql`
      UPDATE vendors
      SET
        name = ${name ?? current.name},
        trade_type = ${trade_type ?? current.trade_type},
        email = ${email !== undefined ? email : current.email},
        phone = ${phone !== undefined ? phone : current.phone},
        notes = ${notes !== undefined ? notes : current.notes},
        insurance_status = ${insurance_status ?? current.insurance_status},
        availability = ${availability ?? current.availability},
        rating = ${rating ?? current.rating},
        updated_at = NOW()
      WHERE id = ${id}
        AND organization_id = ${user.organization_id}
      RETURNING *
    `

    return NextResponse.json(updated[0])
  } catch (error) {
    console.error('Error updating vendor:', error)
    return NextResponse.json({ error: 'Internal server error' }, { status: 500 })
  }
}

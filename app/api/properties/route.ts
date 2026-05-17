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

    const properties = await sql`
      SELECT p.*, COUNT(u.id)::int AS unit_count
      FROM properties p
      LEFT JOIN units u ON u.property_id = p.id
      WHERE p.organization_id = ${user.organization_id}
      GROUP BY p.id
      ORDER BY p.name ASC
    `

    return NextResponse.json(properties)
  } catch (error) {
    console.error('Error fetching properties:', error)
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
      address,
      city = null,
      province = null,
    } = body

    if (!name || !address) {
      return NextResponse.json({ error: 'Name and address are required' }, { status: 400 })
    }

    const inserted = await sql`
      INSERT INTO properties (organization_id, name, address, city, province)
      VALUES (${user.organization_id}, ${name}, ${address}, ${city}, ${province})
      RETURNING *
    `

    return NextResponse.json(inserted[0], { status: 201 })
  } catch (error) {
    console.error('Error creating property:', error)
    return NextResponse.json({ error: 'Internal server error' }, { status: 500 })
  }
}

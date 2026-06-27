import { NextRequest, NextResponse } from 'next/server'
import { getSession } from '@/lib/auth'
import { sql, type Unit } from '@/lib/db'
import { expandUnitSpec } from '@/lib/units'

async function ownsProperty(orgId: string, propertyId: string): Promise<boolean> {
  const rows = (await sql`
    SELECT id FROM properties WHERE id = ${propertyId} AND organization_id = ${orgId}
  `) as { id: string }[]
  return rows.length > 0
}

export async function GET(
  _request: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const session = await getSession()
    if (!session?.user.organization_id) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
    }
    const { id } = await params
    if (!(await ownsProperty(session.user.organization_id, id))) {
      return NextResponse.json({ error: 'Not found' }, { status: 404 })
    }

    const units = (await sql`
      SELECT * FROM units WHERE property_id = ${id} ORDER BY unit_number ASC
    `) as Unit[]
    return NextResponse.json(units)
  } catch (error) {
    console.error('Error listing units:', error)
    return NextResponse.json({ error: 'Internal server error' }, { status: 500 })
  }
}

export async function POST(
  request: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const session = await getSession()
    if (!session?.user.organization_id) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
    }
    const { id } = await params
    if (!(await ownsProperty(session.user.organization_id, id))) {
      return NextResponse.json({ error: 'Not found' }, { status: 404 })
    }

    const body = await request.json()
    let requested: string[] = []
    if (Array.isArray(body?.units)) {
      requested = body.units.map((u: unknown) => String(u).trim()).filter(Boolean)
    } else if (typeof body?.unit_spec === 'string') {
      requested = expandUnitSpec(body.unit_spec)
    }

    if (requested.length === 0) {
      return NextResponse.json({ error: 'No unit numbers provided' }, { status: 400 })
    }

    // Skip units that already exist on this property (case-insensitive) and de-dupe the batch.
    const existing = (await sql`
      SELECT unit_number FROM units WHERE property_id = ${id}
    `) as { unit_number: string }[]
    const have = new Set(existing.map(e => e.unit_number.toLowerCase()))
    const seen = new Set<string>()
    const toAdd = requested.filter(u => {
      const key = u.toLowerCase()
      if (have.has(key) || seen.has(key)) return false
      seen.add(key)
      return true
    })

    if (toAdd.length > 0) {
      await sql`
        INSERT INTO units (property_id, unit_number)
        SELECT ${id}, u FROM unnest(${toAdd}::text[]) AS u
      `
      await sql`
        UPDATE properties
        SET unit_count = (SELECT COUNT(*) FROM units WHERE property_id = ${id}), updated_at = NOW()
        WHERE id = ${id}
      `
    }

    return NextResponse.json(
      { added: toAdd.length, skipped: requested.length - toAdd.length },
      { status: 201 }
    )
  } catch (error) {
    console.error('Error adding units:', error)
    return NextResponse.json({ error: 'Internal server error' }, { status: 500 })
  }
}

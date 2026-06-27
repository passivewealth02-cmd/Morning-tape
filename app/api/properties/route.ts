import { NextRequest, NextResponse } from 'next/server'
import { getSession } from '@/lib/auth'
import { sql, type Organization } from '@/lib/db'
import { checkResourceLimit, getEffectivePlan, getUsage } from '@/lib/plans'
import { expandUnitSpec } from '@/lib/units'

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

type PropertyInput = {
  name?: unknown
  address?: unknown
  city?: unknown
  province?: unknown
  unit_spec?: unknown
  units?: unknown
}

type CleanProperty = {
  index: number
  name: string
  address: string
  city: string | null
  province: string | null
  units: string[]
}

function resolveUnits(input: PropertyInput): string[] {
  if (Array.isArray(input.units)) {
    const seen = new Set<string>()
    const out: string[] = []
    for (const raw of input.units) {
      const t = String(raw).trim()
      if (!t) continue
      const key = t.toLowerCase()
      if (seen.has(key)) continue
      seen.add(key)
      out.push(t)
    }
    return out
  }
  if (typeof input.unit_spec === 'string') return expandUnitSpec(input.unit_spec)
  return []
}

function clean(input: PropertyInput, index: number): CleanProperty {
  const name = typeof input.name === 'string' ? input.name.trim() : ''
  const address = typeof input.address === 'string' ? input.address.trim() : ''
  const city = typeof input.city === 'string' && input.city.trim() ? input.city.trim() : null
  const province = typeof input.province === 'string' && input.province.trim() ? input.province.trim() : null
  return { index, name, address, city, province, units: resolveUnits(input) }
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
    const rawItems: PropertyInput[] = Array.isArray(body?.properties) ? body.properties : [body]
    const items = rawItems.map(clean)

    if (items.length === 0) {
      return NextResponse.json({ error: 'No properties provided' }, { status: 400 })
    }
    const invalid = items.find(i => !i.name || !i.address)
    if (invalid) {
      return NextResponse.json(
        { error: `Each property needs a name and an address (check row ${invalid.index + 1}).` },
        { status: 400 }
      )
    }

    // Plan limit applies to the resulting total number of properties.
    const orgRows = (await sql`
      SELECT * FROM organizations WHERE id = ${user.organization_id}
    `) as unknown as Organization[]
    const plan = getEffectivePlan(orgRows[0])
    const usage = await getUsage(user.organization_id)
    const check = checkResourceLimit(plan, 'properties', usage.properties + items.length - 1)
    if (!check.allowed) {
      const maxAllowed = check.limit === -1 ? Infinity : Math.floor(check.limit * 1.1)
      const remaining = Number.isFinite(maxAllowed) ? Math.max(0, maxAllowed - usage.properties) : items.length
      return NextResponse.json(
        {
          error:
            `Adding ${items.length} ${items.length === 1 ? 'property' : 'properties'} would exceed your plan limit ` +
            `(${usage.properties}/${check.limit}). You can add ${remaining} more — upgrade to ${check.upgrade_to ?? 'a higher plan'} for additional capacity.`,
          limit_exceeded: true,
          upgrade_to: check.upgrade_to,
        },
        { status: 402 }
      )
    }

    const created: { id: string; name: string; units: number }[] = []
    let totalUnits = 0

    for (const item of items) {
      const rows = (await sql`
        INSERT INTO properties (organization_id, name, address, city, province)
        VALUES (${user.organization_id}, ${item.name}, ${item.address}, ${item.city}, ${item.province})
        RETURNING id, name
      `) as { id: string; name: string }[]
      const property = rows[0]

      if (item.units.length > 0) {
        await sql`
          INSERT INTO units (property_id, unit_number)
          SELECT ${property.id}, u FROM unnest(${item.units}::text[]) AS u
        `
        await sql`
          UPDATE properties
          SET unit_count = ${item.units.length}, updated_at = NOW()
          WHERE id = ${property.id}
        `
        totalUnits += item.units.length
      }

      created.push({ id: property.id, name: property.name, units: item.units.length })
    }

    return NextResponse.json(
      {
        created_count: created.length,
        units_count: totalUnits,
        properties: created,
        ...(created.length === 1 ? { id: created[0].id } : {}),
      },
      { status: 201 }
    )
  } catch (error) {
    console.error('Error creating properties:', error)
    return NextResponse.json({ error: 'Internal server error' }, { status: 500 })
  }
}

import { NextRequest, NextResponse } from 'next/server'
import { getSession } from '@/lib/auth'
import { sql } from '@/lib/db'

export async function DELETE(
  _request: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const session = await getSession()
    if (!session?.user.organization_id) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
    }
    const orgId = session.user.organization_id
    const { id } = await params

    const existing = (await sql`
      SELECT id FROM properties WHERE id = ${id} AND organization_id = ${orgId}
    `) as unknown as { id: string }[]
    if (existing.length === 0) {
      return NextResponse.json({ error: 'Property not found' }, { status: 404 })
    }

    // Detach tickets from this property and its units first (FKs have no
    // cascade); the property's units are removed by ON DELETE CASCADE.
    await sql`
      UPDATE maintenance_tickets SET property_id = NULL, unit_id = NULL, updated_at = NOW()
      WHERE property_id = ${id} AND organization_id = ${orgId}
    `
    await sql`DELETE FROM properties WHERE id = ${id} AND organization_id = ${orgId}`

    return NextResponse.json({ deleted: true })
  } catch (error) {
    console.error('Error deleting property:', error)
    return NextResponse.json({ error: 'Internal server error' }, { status: 500 })
  }
}

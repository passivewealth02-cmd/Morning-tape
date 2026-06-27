import { NextRequest, NextResponse } from 'next/server'
import { getSession } from '@/lib/auth'
import { sql } from '@/lib/db'

export async function DELETE(
  _request: NextRequest,
  { params }: { params: Promise<{ id: string; unitId: string }> }
) {
  try {
    const session = await getSession()
    if (!session?.user.organization_id) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
    }
    const { id, unitId } = await params

    // Ensure the property belongs to the caller's organization before deleting.
    const owned = (await sql`
      SELECT 1 FROM properties WHERE id = ${id} AND organization_id = ${session.user.organization_id}
    `) as unknown[]
    if (owned.length === 0) {
      return NextResponse.json({ error: 'Not found' }, { status: 404 })
    }

    await sql`DELETE FROM units WHERE id = ${unitId} AND property_id = ${id}`
    await sql`
      UPDATE properties
      SET unit_count = (SELECT COUNT(*) FROM units WHERE property_id = ${id}), updated_at = NOW()
      WHERE id = ${id}
    `

    return NextResponse.json({ deleted: true })
  } catch (error) {
    console.error('Error deleting unit:', error)
    return NextResponse.json({ error: 'Internal server error' }, { status: 500 })
  }
}

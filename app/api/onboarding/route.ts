import { NextRequest, NextResponse } from 'next/server'
import { getSession } from '@/lib/auth'
import { sql } from '@/lib/db'

function slugify(name: string): string {
  return name
    .toLowerCase()
    .trim()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
}

export async function POST(request: NextRequest) {
  try {
    const session = await getSession()
    if (!session) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
    }

    const { user } = session
    const body = await request.json()
    const { organization_name } = body

    if (!organization_name || typeof organization_name !== 'string') {
      return NextResponse.json({ error: 'Organization name is required' }, { status: 400 })
    }

    const baseSlug = slugify(organization_name)
    const suffix = Math.random().toString(36).slice(2, 6)
    const slug = `${baseSlug}-${suffix}`

    const org = await sql`
      INSERT INTO organizations (name, slug)
      VALUES (${organization_name.trim()}, ${slug})
      RETURNING *
    `

    await sql`
      UPDATE users
      SET organization_id = ${org[0].id}, updated_at = NOW()
      WHERE id = ${user.id}
    `

    return NextResponse.json({ success: true, organization: org[0] })
  } catch (error) {
    console.error('Error during onboarding:', error)
    return NextResponse.json({ error: 'Internal server error' }, { status: 500 })
  }
}

import { NextRequest, NextResponse } from 'next/server'
import { put } from '@vercel/blob'
import { getSession, logActivity } from '@/lib/auth'
import { sql, type TicketFile } from '@/lib/db'

const MAX_SIZE = 4.5 * 1024 * 1024 // 4.5 MB — Vercel serverless request body limit
const ALLOWED_TYPES = [
  'image/jpeg',
  'image/png',
  'image/webp',
  'image/gif',
  'image/heic',
  'application/pdf',
]

export async function POST(
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

    const ticket = (await sql`
      SELECT id FROM maintenance_tickets
      WHERE id = ${id} AND organization_id = ${user.organization_id}
    `) as unknown as { id: string }[]

    if (ticket.length === 0) {
      return NextResponse.json({ error: 'Ticket not found' }, { status: 404 })
    }

    const formData = await request.formData()
    const file = formData.get('file')

    if (!file || !(file instanceof File)) {
      return NextResponse.json({ error: 'No file provided' }, { status: 400 })
    }

    if (!ALLOWED_TYPES.includes(file.type)) {
      return NextResponse.json(
        { error: 'Unsupported file type. Upload an image or PDF.' },
        { status: 400 }
      )
    }

    if (file.size > MAX_SIZE) {
      return NextResponse.json(
        { error: 'File is too large. Maximum size is 4.5 MB.' },
        { status: 400 }
      )
    }

    const blob = await put(`tickets/${id}/${file.name}`, file, {
      access: 'public',
      addRandomSuffix: true,
    })

    const inserted = (await sql`
      INSERT INTO ticket_files (ticket_id, file_url, file_name, file_type, uploaded_by)
      VALUES (${id}, ${blob.url}, ${file.name}, ${file.type}, ${user.id})
      RETURNING *
    `) as unknown as TicketFile[]

    await logActivity({
      organizationId: user.organization_id,
      ticketId: id,
      userId: user.id,
      actionType: 'file_uploaded',
      description: `File uploaded: ${file.name}`,
      metadata: { file_name: file.name, file_type: file.type },
    })

    return NextResponse.json(inserted[0], { status: 201 })
  } catch (error) {
    console.error('Error uploading file:', error)
    return NextResponse.json({ error: 'Internal server error' }, { status: 500 })
  }
}

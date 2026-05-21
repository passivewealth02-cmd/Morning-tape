import { NextRequest, NextResponse } from 'next/server'
import { del, get } from '@vercel/blob'
import { getSession, logActivity } from '@/lib/auth'
import { sql, type TicketFile } from '@/lib/db'

export async function GET(
  _request: NextRequest,
  { params }: { params: Promise<{ id: string; fileId: string }> }
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

    const { id, fileId } = await params

    // Only serve the file if its ticket belongs to the caller's org.
    const rows = (await sql`
      SELECT f.file_url, f.file_name, f.file_type
      FROM ticket_files f
      JOIN maintenance_tickets t ON t.id = f.ticket_id
      WHERE f.id = ${fileId}
        AND f.ticket_id = ${id}
        AND t.organization_id = ${user.organization_id}
    `) as unknown as Pick<TicketFile, 'file_url' | 'file_name' | 'file_type'>[]

    if (rows.length === 0) {
      return NextResponse.json({ error: 'File not found' }, { status: 404 })
    }

    const file = rows[0]
    const result = await get(file.file_url, { access: 'private' })
    if (!result || result.statusCode !== 200) {
      return NextResponse.json({ error: 'File not found' }, { status: 404 })
    }

    const contentType = file.file_type || result.blob.contentType || 'application/octet-stream'
    const safeName = (file.file_name || 'file').replace(/[\r\n"]/g, '')
    return new Response(result.stream, {
      headers: {
        'Content-Type': contentType,
        'Content-Disposition': `inline; filename="${safeName}"`,
        'Cache-Control': 'private, max-age=3600',
      },
    })
  } catch (error) {
    console.error('Error fetching file:', error)
    return NextResponse.json({ error: 'Internal server error' }, { status: 500 })
  }
}

export async function DELETE(
  _request: NextRequest,
  { params }: { params: Promise<{ id: string; fileId: string }> }
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

    const { id, fileId } = await params

    // Only return the file if its ticket belongs to the caller's org
    const rows = (await sql`
      SELECT f.id, f.file_url
      FROM ticket_files f
      JOIN maintenance_tickets t ON t.id = f.ticket_id
      WHERE f.id = ${fileId}
        AND f.ticket_id = ${id}
        AND t.organization_id = ${user.organization_id}
    `) as unknown as Pick<TicketFile, 'id' | 'file_url'>[]

    if (rows.length === 0) {
      return NextResponse.json({ error: 'File not found' }, { status: 404 })
    }

    try {
      await del(rows[0].file_url)
    } catch (blobError) {
      // If the blob is already gone, still remove the DB record
      console.error('Blob delete failed:', blobError)
    }

    await sql`DELETE FROM ticket_files WHERE id = ${fileId}`

    await logActivity({
      organizationId: user.organization_id,
      ticketId: id,
      userId: user.id,
      actionType: 'file_deleted',
      description: 'File removed',
      metadata: { file_id: fileId },
    })

    return NextResponse.json({ success: true })
  } catch (error) {
    console.error('Error deleting file:', error)
    return NextResponse.json({ error: 'Internal server error' }, { status: 500 })
  }
}

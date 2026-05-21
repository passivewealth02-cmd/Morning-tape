import { NextRequest, NextResponse } from 'next/server'
import { put } from '@vercel/blob'
import { getSession, logActivity } from '@/lib/auth'
import { sql, type TicketFile } from '@/lib/db'

const MAX_SIZE = 4.5 * 1024 * 1024 // 4.5 MB — Vercel serverless request body limit
const MAX_FILES_PER_TICKET = 20
const ALLOWED_EXTENSIONS = new Set(['jpg', 'jpeg', 'png', 'webp', 'gif', 'heic', 'heif', 'pdf'])

// Detect the real file type from its leading bytes. The client-supplied MIME type
// is untrusted (trivially spoofable), so an attacker could otherwise upload active
// content (HTML/JS/SVG) declared as image/png. We only accept files whose actual
// bytes match a supported image or PDF signature.
function sniffMime(bytes: Uint8Array): string | null {
  const at = (sig: number[], offset = 0) => sig.every((b, i) => bytes[offset + i] === b)

  if (at([0xff, 0xd8, 0xff])) return 'image/jpeg'
  if (at([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) return 'image/png'
  if (at([0x47, 0x49, 0x46, 0x38])) return 'image/gif'
  if (at([0x52, 0x49, 0x46, 0x46]) && at([0x57, 0x45, 0x42, 0x50], 8)) return 'image/webp'
  if (at([0x25, 0x50, 0x44, 0x46])) return 'application/pdf'
  if (at([0x66, 0x74, 0x79, 0x70], 4)) {
    const brand = String.fromCharCode(...bytes.slice(8, 12))
    if (['heic', 'heix', 'hevc', 'hevx', 'mif1', 'msf1', 'heim', 'heis'].includes(brand)) {
      return 'image/heic'
    }
  }
  return null
}

function sanitizeFilename(name: string): string {
  const base = name.split(/[\\/]/).pop() || 'file'
  return base.replace(/[^a-zA-Z0-9._-]/g, '_').slice(0, 200) || 'file'
}

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

    const countRows = (await sql`
      SELECT COUNT(*)::int AS count FROM ticket_files WHERE ticket_id = ${id}
    `) as unknown as { count: number }[]
    if ((countRows[0]?.count ?? 0) >= MAX_FILES_PER_TICKET) {
      return NextResponse.json(
        { error: `This ticket already has the maximum of ${MAX_FILES_PER_TICKET} files.` },
        { status: 400 }
      )
    }

    const formData = await request.formData()
    const file = formData.get('file')

    if (!file || !(file instanceof File)) {
      return NextResponse.json({ error: 'No file provided' }, { status: 400 })
    }

    if (file.size > MAX_SIZE) {
      return NextResponse.json(
        { error: 'File is too large. Maximum size is 4.5 MB.' },
        { status: 400 }
      )
    }

    const safeName = sanitizeFilename(file.name)
    const ext = safeName.includes('.') ? safeName.split('.').pop()!.toLowerCase() : ''
    if (!ALLOWED_EXTENSIONS.has(ext)) {
      return NextResponse.json(
        { error: 'Unsupported file type. Upload an image or PDF.' },
        { status: 400 }
      )
    }

    const buffer = await file.arrayBuffer()
    const detectedType = sniffMime(new Uint8Array(buffer))
    if (!detectedType) {
      return NextResponse.json(
        { error: 'File contents do not match a supported image or PDF.' },
        { status: 400 }
      )
    }

    // Store privately — attachments are served only through the authenticated
    // download proxy (GET on this file's route), never via a public blob URL.
    const blob = await put(`tickets/${id}/${safeName}`, buffer, {
      access: 'private',
      addRandomSuffix: true,
      contentType: detectedType,
    })

    const inserted = (await sql`
      INSERT INTO ticket_files (ticket_id, file_url, file_name, file_type, uploaded_by)
      VALUES (${id}, ${blob.url}, ${safeName}, ${detectedType}, ${user.id})
      RETURNING *
    `) as unknown as TicketFile[]

    await logActivity({
      organizationId: user.organization_id,
      ticketId: id,
      userId: user.id,
      actionType: 'file_uploaded',
      description: `File uploaded: ${safeName}`,
      metadata: { file_name: safeName, file_type: detectedType },
    })

    return NextResponse.json(inserted[0], { status: 201 })
  } catch (error) {
    console.error('Error uploading file:', error)
    // TEMPORARY: surface the underlying error to diagnose upload failures.
    const detail = error instanceof Error ? error.message : String(error)
    return NextResponse.json({ error: `Upload failed: ${detail}` }, { status: 500 })
  }
}

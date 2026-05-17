import { NextRequest, NextResponse } from 'next/server'
import { getSession, logActivity } from '@/lib/auth'
import { sql } from '@/lib/db'

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

    const ticket = await sql`
      SELECT id FROM maintenance_tickets
      WHERE id = ${id} AND organization_id = ${user.organization_id}
    `

    if (ticket.length === 0) {
      return NextResponse.json({ error: 'Ticket not found' }, { status: 404 })
    }

    const body = await request.json()
    const {
      message,
      is_internal = false,
      sender_type = 'manager',
    } = body

    if (!message) {
      return NextResponse.json({ error: 'Message is required' }, { status: 400 })
    }

    const inserted = await sql`
      INSERT INTO ticket_messages (ticket_id, sender_type, sender_id, message, is_internal)
      VALUES (${id}, ${sender_type}, ${user.id}, ${message}, ${is_internal})
      RETURNING *
    `

    const newMessage = inserted[0]

    await logActivity({
      organizationId: user.organization_id,
      ticketId: id,
      userId: user.id,
      actionType: 'message_added',
      description: is_internal ? 'Internal note added' : 'Message sent',
      metadata: { sender_type, is_internal },
    })

    return NextResponse.json(newMessage, { status: 201 })
  } catch (error) {
    console.error('Error adding message:', error)
    return NextResponse.json({ error: 'Internal server error' }, { status: 500 })
  }
}

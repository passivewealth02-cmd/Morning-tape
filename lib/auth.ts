import 'server-only'
import { cookies } from 'next/headers'
import { sql, type User } from './db'
import { randomBytes } from 'crypto'

export const SESSION_COOKIE_NAME = 'maintena_session'
const SESSION_DURATION_DAYS = 30

export function sessionCookieOptions(expiresAt: Date) {
  return {
    httpOnly: true,
    secure: process.env.NODE_ENV === 'production',
    sameSite: 'lax' as const,
    expires: expiresAt,
    path: '/',
  }
}

export async function createMagicLinkToken(email: string): Promise<string> {
  const token = randomBytes(32).toString('hex')
  const expiresAt = new Date(Date.now() + 15 * 60 * 1000)

  await sql`
    INSERT INTO magic_link_tokens (email, token, expires_at)
    VALUES (${email}, ${token}, ${expiresAt.toISOString()})
  `

  return token
}

export async function verifyMagicLinkToken(token: string): Promise<string | null> {
  const result = await sql`
    SELECT email FROM magic_link_tokens
    WHERE token = ${token}
      AND expires_at > NOW()
      AND used = FALSE
  `

  if (result.length === 0) {
    return null
  }

  await sql`
    UPDATE magic_link_tokens SET used = TRUE WHERE token = ${token}
  `

  return result[0].email
}

export async function getOrCreateUser(email: string, name?: string): Promise<User> {
  const existing = (await sql`SELECT * FROM users WHERE email = ${email}`) as unknown as User[]

  if (existing.length > 0) {
    const user = existing[0]
    if (name && !user.name) {
      await sql`UPDATE users SET name = ${name}, updated_at = NOW() WHERE id = ${user.id}`
      user.name = name
    }
    return user
  }

  const newUser = (await sql`
    INSERT INTO users (email, name, role)
    VALUES (${email}, ${name ?? null}, 'admin')
    RETURNING *
  `) as unknown as User[]

  return newUser[0]
}

export async function createSession(userId: string): Promise<{ token: string; expiresAt: Date }> {
  const token = randomBytes(32).toString('hex')
  const expiresAt = new Date(Date.now() + SESSION_DURATION_DAYS * 24 * 60 * 60 * 1000)

  await sql`
    INSERT INTO sessions (user_id, token, expires_at)
    VALUES (${userId}, ${token}, ${expiresAt.toISOString()})
  `

  return { token, expiresAt }
}

export async function getSession(): Promise<{ user: User } | null> {
  const cookieStore = await cookies()
  const sessionToken = cookieStore.get(SESSION_COOKIE_NAME)?.value

  if (!sessionToken) return null

  const sessionResult = await sql`
    SELECT s.user_id FROM sessions s
    WHERE s.token = ${sessionToken} AND s.expires_at > NOW()
  `

  if (sessionResult.length === 0) return null

  const userResult = await sql`
    SELECT * FROM users WHERE id = ${sessionResult[0].user_id}
  `

  if (userResult.length === 0) return null

  return { user: userResult[0] as User }
}

export async function destroySession(): Promise<void> {
  const cookieStore = await cookies()
  const sessionToken = cookieStore.get(SESSION_COOKIE_NAME)?.value

  if (sessionToken) {
    await sql`DELETE FROM sessions WHERE token = ${sessionToken}`
  }

  cookieStore.delete(SESSION_COOKIE_NAME)
}

export async function requireSession() {
  const session = await getSession()
  if (!session) {
    throw new Error('Unauthorized')
  }
  return session
}

export async function logActivity({
  organizationId,
  ticketId,
  userId,
  actionType,
  description,
  metadata,
}: {
  organizationId: string
  ticketId?: string
  userId?: string
  actionType: string
  description: string
  metadata?: Record<string, unknown>
}) {
  await sql`
    INSERT INTO activity_logs (organization_id, ticket_id, user_id, action_type, description, metadata)
    VALUES (
      ${organizationId},
      ${ticketId ?? null},
      ${userId ?? null},
      ${actionType},
      ${description},
      ${metadata ? JSON.stringify(metadata) : null}
    )
  `
}

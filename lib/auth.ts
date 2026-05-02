import 'server-only'
import { cookies } from 'next/headers'
import { sql, type User, type Subscription } from './db'
import { randomBytes } from 'crypto'

const SESSION_COOKIE_NAME = 'morning_tape_session'
const SESSION_DURATION_DAYS = 30

export async function createMagicLinkToken(email: string): Promise<string> {
  const token = randomBytes(32).toString('hex')
  const expiresAt = new Date(Date.now() + 15 * 60 * 1000) // 15 minutes

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

  // Mark token as used
  await sql`
    UPDATE magic_link_tokens
    SET used = TRUE
    WHERE token = ${token}
  `

  return result[0].email
}

export async function getOrCreateUser(email: string): Promise<User> {
  // Try to find existing user
  const existingUser = await sql`
    SELECT * FROM users WHERE email = ${email}
  `

  if (existingUser.length > 0) {
    return existingUser[0] as User
  }

  // Create new user
  const newUser = await sql`
    INSERT INTO users (email)
    VALUES (${email})
    RETURNING *
  `

  return newUser[0] as User
}

export async function createSession(userId: string): Promise<string> {
  const token = randomBytes(32).toString('hex')
  const expiresAt = new Date(Date.now() + SESSION_DURATION_DAYS * 24 * 60 * 60 * 1000)

  await sql`
    INSERT INTO sessions (user_id, token, expires_at)
    VALUES (${userId}, ${token}, ${expiresAt.toISOString()})
  `

  // Set HTTP-only cookie
  const cookieStore = await cookies()
  cookieStore.set(SESSION_COOKIE_NAME, token, {
    httpOnly: true,
    secure: process.env.NODE_ENV === 'production',
    sameSite: 'lax',
    expires: expiresAt,
    path: '/',
  })

  return token
}

export async function getSession(): Promise<{ user: User; subscription: Subscription | null } | null> {
  const cookieStore = await cookies()
  const sessionToken = cookieStore.get(SESSION_COOKIE_NAME)?.value

  if (!sessionToken) {
    return null
  }

  const sessionResult = await sql`
    SELECT s.user_id, s.expires_at
    FROM sessions s
    WHERE s.token = ${sessionToken}
      AND s.expires_at > NOW()
  `

  if (sessionResult.length === 0) {
    return null
  }

  const userId = sessionResult[0].user_id

  const userResult = await sql`
    SELECT * FROM users WHERE id = ${userId}
  `

  if (userResult.length === 0) {
    return null
  }

  const subscriptionResult = await sql`
    SELECT * FROM subscriptions
    WHERE user_id = ${userId}
      AND status = 'active'
    ORDER BY created_at DESC
    LIMIT 1
  `

  return {
    user: userResult[0] as User,
    subscription: subscriptionResult.length > 0 ? (subscriptionResult[0] as Subscription) : null,
  }
}

export async function destroySession(): Promise<void> {
  const cookieStore = await cookies()
  const sessionToken = cookieStore.get(SESSION_COOKIE_NAME)?.value

  if (sessionToken) {
    await sql`
      DELETE FROM sessions WHERE token = ${sessionToken}
    `
  }

  cookieStore.delete(SESSION_COOKIE_NAME)
}

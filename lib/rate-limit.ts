import 'server-only'
import { sql } from './db'

// Fixed-window rate limiter backed by Postgres (no extra infra required).
// Each (bucket, window) pair counts requests; once the count exceeds the
// limit within the window, further requests are rejected until the window rolls.

let ensured = false

async function ensureTable() {
  if (ensured) return
  await sql`
    CREATE TABLE IF NOT EXISTS rate_limits (
      bucket TEXT NOT NULL,
      window_start TIMESTAMPTZ NOT NULL,
      count INTEGER NOT NULL DEFAULT 0,
      PRIMARY KEY (bucket, window_start)
    )
  `
  // Opportunistic cleanup of stale windows so the table can't grow unbounded.
  await sql`DELETE FROM rate_limits WHERE window_start < NOW() - INTERVAL '1 day'`
  ensured = true
}

export type RateLimitResult = { allowed: boolean; remaining: number; retryAfter: number }

/**
 * Returns whether a request for `key` is allowed within `limit` per `windowSeconds`.
 * Fails open (allows) if the rate-limit store is unreachable, so a DB hiccup never
 * locks legitimate users out of public intake.
 */
export async function rateLimit(
  key: string,
  limit: number,
  windowSeconds: number
): Promise<RateLimitResult> {
  try {
    await ensureTable()
    const nowMs = Date.now()
    const windowMs = windowSeconds * 1000
    const windowStart = new Date(Math.floor(nowMs / windowMs) * windowMs)

    const rows = (await sql`
      INSERT INTO rate_limits (bucket, window_start, count)
      VALUES (${key}, ${windowStart.toISOString()}, 1)
      ON CONFLICT (bucket, window_start)
      DO UPDATE SET count = rate_limits.count + 1
      RETURNING count
    `) as unknown as { count: number }[]

    const count = rows[0]?.count ?? 1
    const allowed = count <= limit
    const retryAfter = allowed
      ? 0
      : Math.ceil((windowStart.getTime() + windowMs - nowMs) / 1000)
    return { allowed, remaining: Math.max(0, limit - count), retryAfter }
  } catch (err) {
    console.error('Rate limiter unavailable, allowing request:', err)
    return { allowed: true, remaining: 0, retryAfter: 0 }
  }
}

export function clientIp(request: Request): string {
  const xff = request.headers.get('x-forwarded-for')
  if (xff) return xff.split(',')[0]!.trim()
  return request.headers.get('x-real-ip') ?? 'unknown'
}

export function tooManyRequests(retryAfter: number) {
  return new Response(
    JSON.stringify({ error: 'Too many requests. Please slow down and try again shortly.' }),
    {
      status: 429,
      headers: {
        'Content-Type': 'application/json',
        'Retry-After': String(Math.max(1, retryAfter)),
      },
    }
  )
}

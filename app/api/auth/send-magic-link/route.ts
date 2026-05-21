import { NextRequest, NextResponse } from 'next/server'
import { createMagicLinkToken } from '@/lib/auth'
import { sendMagicLinkEmail } from '@/lib/email'
import { rateLimit, clientIp, tooManyRequests } from '@/lib/rate-limit'

export async function POST(request: NextRequest) {
  try {
    // Limit per IP to prevent abuse of the email sender across many addresses.
    const ipLimit = await rateLimit(`magic:ip:${clientIp(request)}`, 10, 3600)
    if (!ipLimit.allowed) return tooManyRequests(ipLimit.retryAfter)

    const { email } = await request.json()

    if (!email || typeof email !== 'string') {
      return NextResponse.json(
        { error: 'Email is required' },
        { status: 400 }
      )
    }

    // Basic email validation
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/
    if (!emailRegex.test(email)) {
      return NextResponse.json(
        { error: 'Invalid email format' },
        { status: 400 }
      )
    }

    // Limit per address to prevent email-bombing a specific target.
    const emailLimit = await rateLimit(`magic:email:${email.toLowerCase()}`, 5, 900)
    if (!emailLimit.allowed) return tooManyRequests(emailLimit.retryAfter)

    // Create magic link token
    const token = await createMagicLinkToken(email.toLowerCase())

    // Derive base URL from request so magic links work on any deployment
    const baseUrl = `${request.nextUrl.protocol}//${request.nextUrl.host}`

    // Send email
    const sent = await sendMagicLinkEmail(email.toLowerCase(), token, baseUrl)

    if (!sent) {
      return NextResponse.json(
        { error: 'Failed to send email. Please try again.' },
        { status: 500 }
      )
    }

    return NextResponse.json({ success: true })
  } catch (error) {
    console.error('Error sending magic link:', error)
    return NextResponse.json(
      { error: 'An error occurred. Please try again.' },
      { status: 500 }
    )
  }
}

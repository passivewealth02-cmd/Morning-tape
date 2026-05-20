import { NextRequest, NextResponse } from 'next/server'
import { verifyMagicLinkToken, getOrCreateUser, createSession } from '@/lib/auth'

export async function GET(request: NextRequest) {
  const token = request.nextUrl.searchParams.get('token')
  const origin = `${request.nextUrl.protocol}//${request.nextUrl.host}`

  if (!token) {
    return NextResponse.redirect(`${origin}/login?error=missing_token`)
  }

  const email = await verifyMagicLinkToken(token)
  if (!email) {
    return NextResponse.redirect(`${origin}/login?error=invalid_token`)
  }

  const user = await getOrCreateUser(email)
  await createSession(user.id)

  const destination = user.organization_id ? '/dashboard' : '/onboarding'
  return NextResponse.redirect(`${origin}${destination}`)
}

import { NextRequest, NextResponse } from 'next/server'
import { cookies } from 'next/headers'
import { getOrCreateUser, createSession } from '@/lib/auth'

export async function GET(request: NextRequest) {
  const origin = `${request.nextUrl.protocol}//${request.nextUrl.host}`
  const { searchParams } = request.nextUrl
  const code = searchParams.get('code')
  const state = searchParams.get('state')
  const errorParam = searchParams.get('error')

  if (errorParam) {
    return NextResponse.redirect(`${origin}/login?error=google_denied`)
  }

  const cookieStore = await cookies()
  const savedState = cookieStore.get('oauth_state')?.value
  cookieStore.delete('oauth_state')

  if (!code || !state || state !== savedState) {
    return NextResponse.redirect(`${origin}/login?error=oauth_state_mismatch`)
  }

  const baseUrl = process.env.NEXT_PUBLIC_APP_URL?.replace(/\/$/, '') || origin
  const redirectUri = `${baseUrl}/api/auth/callback/google`

  const tokenRes = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      code,
      client_id: process.env.GOOGLE_CLIENT_ID!,
      client_secret: process.env.GOOGLE_CLIENT_SECRET!,
      redirect_uri: redirectUri,
      grant_type: 'authorization_code',
    }),
  })

  if (!tokenRes.ok) {
    console.error('Google token exchange failed:', await tokenRes.text())
    return NextResponse.redirect(`${origin}/login?error=google_token_failed`)
  }

  const { access_token } = await tokenRes.json()

  const userInfoRes = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
    headers: { Authorization: `Bearer ${access_token}` },
  })

  if (!userInfoRes.ok) {
    return NextResponse.redirect(`${origin}/login?error=google_userinfo_failed`)
  }

  const { email, name } = await userInfoRes.json()

  if (!email) {
    return NextResponse.redirect(`${origin}/login?error=google_no_email`)
  }

  try {
    const user = await getOrCreateUser(email.toLowerCase(), name ?? undefined)
    await createSession(user.id)
    const destination = user.organization_id ? '/dashboard' : '/onboarding'
    return NextResponse.redirect(`${origin}${destination}`)
  } catch (error) {
    console.error('OAuth sign-in error:', error)
    return NextResponse.redirect(`${origin}/login?error=signin_failed`)
  }
}

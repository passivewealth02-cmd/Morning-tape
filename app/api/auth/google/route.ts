import { NextRequest, NextResponse } from 'next/server'
import { randomBytes } from 'crypto'
import { cookies } from 'next/headers'

export async function GET(request: NextRequest) {
  if (!process.env.GOOGLE_CLIENT_ID) {
    const origin = `${request.nextUrl.protocol}//${request.nextUrl.host}`
    return NextResponse.redirect(`${origin}/login?error=google_not_configured`)
  }

  const state = randomBytes(16).toString('hex')
  const cookieStore = await cookies()
  cookieStore.set('oauth_state', state, {
    httpOnly: true,
    secure: process.env.NODE_ENV === 'production',
    sameSite: 'lax',
    maxAge: 600,
    path: '/',
  })

  const origin = `${request.nextUrl.protocol}//${request.nextUrl.host}`
  const baseUrl = process.env.NEXT_PUBLIC_APP_URL?.replace(/\/$/, '') || origin
  const redirectUri = `${baseUrl}/api/auth/callback/google`

  const authUrl = new URL('https://accounts.google.com/o/oauth2/v2/auth')
  authUrl.searchParams.set('client_id', process.env.GOOGLE_CLIENT_ID)
  authUrl.searchParams.set('redirect_uri', redirectUri)
  authUrl.searchParams.set('response_type', 'code')
  authUrl.searchParams.set('scope', 'email profile')
  authUrl.searchParams.set('state', state)
  authUrl.searchParams.set('prompt', 'select_account')

  return NextResponse.redirect(authUrl.toString())
}

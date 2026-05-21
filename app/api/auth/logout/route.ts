import { NextRequest, NextResponse } from 'next/server'
import { destroySession, SESSION_COOKIE_NAME } from '@/lib/auth'

export async function POST(request: NextRequest) {
  const origin = `${request.nextUrl.protocol}//${request.nextUrl.host}`
  try {
    await destroySession()
  } catch (error) {
    console.error('Error logging out:', error)
  }
  const response = NextResponse.redirect(`${origin}/login`, { status: 303 })
  response.cookies.delete(SESSION_COOKIE_NAME)
  return response
}

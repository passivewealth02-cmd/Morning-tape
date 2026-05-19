import { NextRequest, NextResponse } from 'next/server'
import { destroySession } from '@/lib/auth'

export async function POST(request: NextRequest) {
  const origin = `${request.nextUrl.protocol}//${request.nextUrl.host}`
  try {
    await destroySession()
  } catch (error) {
    console.error('Error logging out:', error)
  }
  return NextResponse.redirect(`${origin}/login`, { status: 303 })
}

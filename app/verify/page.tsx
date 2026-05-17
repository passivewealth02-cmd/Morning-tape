import { redirect } from 'next/navigation'
import { verifyMagicLinkToken, getOrCreateUser, createSession } from '@/lib/auth'
import Link from 'next/link'

interface VerifyPageProps {
  searchParams: Promise<{ token?: string }>
}

export default async function VerifyPage({ searchParams }: VerifyPageProps) {
  const params = await searchParams
  const { token } = params

  if (!token) {
    return (
      <div className="min-h-screen flex items-center justify-center px-6 bg-gray-50">
        <div className="max-w-md w-full text-center">
          <h1 className="text-2xl font-semibold text-gray-900 mb-3">Invalid link</h1>
          <p className="text-gray-500 mb-6">This sign-in link is invalid or has expired.</p>
          <Link href="/login" className="text-indigo-600 hover:underline text-sm">Request a new link</Link>
        </div>
      </div>
    )
  }

  const email = await verifyMagicLinkToken(token)

  if (!email) {
    return (
      <div className="min-h-screen flex items-center justify-center px-6 bg-gray-50">
        <div className="max-w-md w-full text-center">
          <h1 className="text-2xl font-semibold text-gray-900 mb-3">Link expired</h1>
          <p className="text-gray-500 mb-6">This sign-in link has expired or already been used.</p>
          <Link href="/login" className="text-indigo-600 hover:underline text-sm">Request a new link</Link>
        </div>
      </div>
    )
  }

  const user = await getOrCreateUser(email)
  await createSession(user.id)

  // If user has no organization, redirect to onboarding
  if (!user.organization_id) {
    redirect('/onboarding')
  }

  redirect('/dashboard')
}

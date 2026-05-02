import { redirect } from 'next/navigation'
import { verifyMagicLinkToken, getOrCreateUser, createSession, getSession } from '@/lib/auth'
import { sql } from '@/lib/db'
import Link from 'next/link'

interface VerifyPageProps {
  searchParams: Promise<{ token?: string; plan?: string }>
}

export default async function VerifyPage({ searchParams }: VerifyPageProps) {
  const params = await searchParams
  const { token, plan } = params

  if (!token) {
    return (
      <div className="paper-texture min-h-screen flex items-center justify-center px-6">
        <div className="max-w-md w-full text-center">
          <h1 className="font-serif text-3xl font-semibold mb-4">
            Invalid Link
          </h1>
          <p className="text-muted-foreground mb-6">
            This sign-in link is invalid or has expired.
          </p>
          <Link href="/login" className="editorial-link">
            Request a new link
          </Link>
        </div>
      </div>
    )
  }

  const email = await verifyMagicLinkToken(token)

  if (!email) {
    return (
      <div className="paper-texture min-h-screen flex items-center justify-center px-6">
        <div className="max-w-md w-full text-center">
          <h1 className="font-serif text-3xl font-semibold mb-4">
            Link Expired
          </h1>
          <p className="text-muted-foreground mb-6">
            This sign-in link has expired or has already been used.
          </p>
          <Link href="/login" className="editorial-link">
            Request a new link
          </Link>
        </div>
      </div>
    )
  }

  // Create or get user
  const user = await getOrCreateUser(email)

  // Create session
  await createSession(user.id)

  // Check if user has an active subscription
  const subscriptionResult = await sql`
    SELECT * FROM subscriptions
    WHERE user_id = ${user.id}
      AND status = 'active'
    LIMIT 1
  `

  const hasSubscription = subscriptionResult.length > 0

  // If user has subscription, go to dashboard
  // Otherwise, redirect to subscribe page (optionally with plan)
  if (hasSubscription) {
    redirect('/dashboard')
  } else {
    const planParam = plan ? `?plan=${plan}` : ''
    redirect(`/subscribe${planParam}`)
  }
}

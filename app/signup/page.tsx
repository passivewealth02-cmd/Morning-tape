'use client'

import { useState, Suspense } from 'react'
import { useSearchParams } from 'next/navigation'
import Link from 'next/link'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'

const ERROR_MESSAGES: Record<string, string> = {
  missing_token: 'Sign-up link is missing its token. Request a new one.',
  invalid_token: 'This sign-up link has expired or already been used. Request a new one.',
  google_not_configured: 'Google sign-up is not configured yet.',
  google_denied: 'Google sign-up was cancelled.',
  oauth_state_mismatch: 'Sign-up session expired. Please try again.',
  google_token_failed: 'Could not complete Google sign-up. Please try again.',
  google_userinfo_failed: 'Could not retrieve your Google profile. Please try again.',
  google_no_email: 'Your Google account did not provide an email address.',
  signin_failed: 'Sign-up failed. Please try again.',
}

function GoogleIcon() {
  return (
    <svg className="w-4 h-4" viewBox="0 0 24 24">
      <path fill="#4285F4" d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z"/>
      <path fill="#34A853" d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z"/>
      <path fill="#FBBC05" d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l2.85-2.22.81-.62z"/>
      <path fill="#EA4335" d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z"/>
    </svg>
  )
}

function SignupForm() {
  const searchParams = useSearchParams()
  const linkError = searchParams.get('error')
  const [email, setEmail] = useState('')
  const [status, setStatus] = useState<'idle' | 'loading' | 'success' | 'error'>('idle')
  const [errorMessage, setErrorMessage] = useState(
    linkError && ERROR_MESSAGES[linkError] ? ERROR_MESSAGES[linkError] : ''
  )

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    setStatus('loading')
    try {
      const response = await fetch('/api/auth/send-magic-link', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email }),
      })
      const data = await response.json()
      if (!response.ok) throw new Error(data.error || 'Failed to send sign-up link')
      setStatus('success')
    } catch (error) {
      setStatus('error')
      setErrorMessage(error instanceof Error ? error.message : 'An error occurred')
    }
  }

  if (status === 'success') {
    return (
      <div className="min-h-screen flex items-center justify-center px-6 bg-gray-50">
        <div className="max-w-md w-full text-center">
          <div className="w-12 h-12 bg-green-100 rounded-full flex items-center justify-center mx-auto mb-4">
            <svg className="w-6 h-6 text-green-600" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
            </svg>
          </div>
          <h1 className="text-2xl font-semibold text-gray-900 mb-2">Check your inbox</h1>
          <p className="text-gray-500 mb-6">We sent a sign-up link to <strong>{email}</strong></p>
          <button onClick={() => setStatus('idle')} className="text-sm text-indigo-600 hover:underline">
            Try a different email
          </button>
        </div>
      </div>
    )
  }

  return (
    <div className="min-h-screen flex items-center justify-center px-6 bg-gray-50">
      <div className="max-w-sm w-full">
        <div className="text-center mb-8">
          <Link href="/" className="text-xl font-semibold text-gray-900">Maintena</Link>
          <h1 className="text-gray-900 text-lg font-semibold mt-3">Create your account</h1>
          <p className="text-gray-500 text-sm mt-1">14-day free trial — no credit card required</p>
        </div>

        <div className="bg-white rounded-xl border border-gray-200 p-8 shadow-sm space-y-4">
          <a
            href="/api/auth/google"
            className="flex items-center justify-center gap-2.5 w-full border border-gray-200 rounded-lg px-4 py-2.5 text-sm font-medium text-gray-700 hover:bg-gray-50 transition-colors"
          >
            <GoogleIcon />
            Sign up with Google
          </a>

          <div className="flex items-center gap-3">
            <div className="flex-1 h-px bg-gray-200" />
            <span className="text-xs text-gray-400">or</span>
            <div className="flex-1 h-px bg-gray-200" />
          </div>

          <form onSubmit={handleSubmit} className="space-y-4">
            <div>
              <label htmlFor="email" className="block text-sm font-medium text-gray-700 mb-1.5">
                Work email
              </label>
              <Input
                id="email"
                type="email"
                value={email}
                onChange={e => setEmail(e.target.value)}
                placeholder="you@company.com"
                required
                disabled={status === 'loading'}
                className="w-full"
              />
            </div>

            {(status === 'error' || (status === 'idle' && errorMessage)) && (
              <p className="text-sm text-red-600">{errorMessage}</p>
            )}

            <Button
              type="submit"
              disabled={status === 'loading' || !email}
              className="w-full bg-indigo-600 hover:bg-indigo-700 text-white"
            >
              {status === 'loading' ? 'Sending link...' : 'Create account'}
            </Button>
          </form>

          <p className="text-center text-xs text-gray-400">
            By signing up you agree to our terms.
          </p>
        </div>

        <p className="text-center text-sm text-gray-500 mt-6">
          Already have an account?{' '}
          <Link href="/login" className="text-indigo-600 hover:underline font-medium">
            Sign in
          </Link>
        </p>
      </div>
    </div>
  )
}

export default function SignupPage() {
  return (
    <Suspense fallback={
      <div className="min-h-screen flex items-center justify-center bg-gray-50">
        <div className="w-5 h-5 border-2 border-gray-300 border-t-gray-700 rounded-full animate-spin" />
      </div>
    }>
      <SignupForm />
    </Suspense>
  )
}

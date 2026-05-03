'use client'

import { useState, Suspense } from 'react'
import Link from 'next/link'
import { useSearchParams } from 'next/navigation'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Spinner } from '@/components/ui/spinner'

function LoginForm() {
  const searchParams = useSearchParams()
  const plan = searchParams.get('plan')
  const [email, setEmail] = useState('')
  const [status, setStatus] = useState('idle')
  const [errorMessage, setErrorMessage] = useState('')

  const handleSubmit = async (e) => {
    e.preventDefault()
    setStatus('loading')
    try {
      const response = await fetch('/api/auth/send-magic-link', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email, plan }),
      })
      const data = await response.json()
      if (!response.ok) throw new Error(data.error || 'Failed to send magic link')
      setStatus('success')
    } catch (error) {
      setStatus('error')
      setErrorMessage(error.message || 'An error occurred')
    }
  }

  if (status === 'success') {
    return (
      <div className="paper-texture min-h-screen flex items-center justify-center px-6">
        <div className="max-w-md w-full text-center">
          <h1 className="font-serif text-3xl font-semibold mb-4">Check Your Inbox</h1>
          <p className="text-muted-foreground mb-6">We sent a sign-in link to {email}.</p>
          <button onClick={() => setStatus('idle')} className="editorial-link">Try again</button>
        </div>
      </div>
    )
  }

  return (
    <div className="paper-texture min-h-screen flex items-center justify-center px-6">
      <div className="max-w-md w-full">
        <div className="text-center mb-8">
          <Link href="/"><h1 className="masthead text-2xl font-serif font-semibold tracking-wider mb-2">The Morning Tape</h1></Link>
          <p className="section-marker">— Subscriber Access —</p>
        </div>
        <form onSubmit={handleSubmit} className="space-y-6">
          <Input id="email" type="email" value={email} onChange={(e) => setEmail(e.target.value)} placeholder="reader@example.com" required disabled={status === 'loading'} />
          {status === 'error' && <p className="text-sm text-accent">{errorMessage}</p>}
          <Button type="submit" disabled={status === 'loading' || !email} className="w-full font-serif tracking-wide bg-foreground text-background">
            {status === 'loading' ? 'Sending...' : 'Send Sign-In Link'}
          </Button>
        </form>
      </div>
    </div>
  )
}

export default function LoginPage() {
  return (
    <Suspense fallback={<div className="paper-texture min-h-screen flex items-center justify-center"><p>Loading...</p></div>}>
      <LoginForm />
    </Suspense>
  )
}

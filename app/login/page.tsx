'use client'

import { useState } from 'react'
import Link from 'next/link'
import { useSearchParams } from 'next/navigation'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Spinner } from '@/components/ui/spinner'

export default function LoginPage() {
  const searchParams = useSearchParams()
  const plan = searchParams.get('plan')
  
  const [email, setEmail] = useState('')
  const [status, setStatus] = useState<'idle' | 'loading' | 'success' | 'error'>('idle')
  const [errorMessage, setErrorMessage] = useState('')

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    setStatus('loading')
    setErrorMessage('')

    try {
      const response = await fetch('/api/auth/send-magic-link', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email, plan }),
      })

      const data = await response.json()

      if (!response.ok) {
        throw new Error(data.error || 'Failed to send magic link')
      }

      setStatus('success')
    } catch (error) {
      setStatus('error')
      setErrorMessage(error instanceof Error ? error.message : 'An error occurred')
    }
  }

  if (status === 'success') {
    return (
      <div className="paper-texture min-h-screen flex items-center justify-center px-6">
        <div className="max-w-md w-full text-center">
          <div className="rule-double pb-6 mb-6" />
          
          <h1 className="font-serif text-3xl font-semibold mb-4">
            Check Your Inbox
          </h1>
          
          <p className="text-muted-foreground mb-6">
            We&apos;ve sent a sign-in link to <span className="font-medium text-foreground">{email}</span>. 
            The link will expire in 15 minutes.
          </p>
          
          <div className="rule-single my-8" />
          
          <p className="text-sm text-muted-foreground">
            Didn&apos;t receive the email? Check your spam folder or{' '}
            <button 
              onClick={() => setStatus('idle')}
              className="editorial-link"
            >
              try again
            </button>
          </p>
        </div>
      </div>
    )
  }

  return (
    <div className="paper-texture min-h-screen flex items-center justify-center px-6">
      <div className="max-w-md w-full">
        <div className="text-center mb-8">
          <Link href="/" className="inline-block">
            <h1 className="masthead text-2xl font-serif font-semibold tracking-wider mb-2">
              The Morning Tape
            </h1>
          </Link>
          <div className="rule-double pb-4 mb-4" />
          <p className="section-marker">— Subscriber Access —</p>
        </div>
        
        <form onSubmit={handleSubmit} className="space-y-6">
          <div>
            <label htmlFor="email" className="block text-sm font-medium mb-2">
              Email Address
            </label>
            <Input
              id="email"
              type="email"
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              placeholder="reader@example.com"
              required
              className="bg-background border-border focus:border-accent focus:ring-accent"
              disabled={status === 'loading'}
            />
          </div>
          
          {status === 'error' && (
            <p className="text-sm text-accent">{errorMessage}</p>
          )}
          
          <Button
            type="submit"
            disabled={status === 'loading' || !email}
            className="w-full font-serif tracking-wide bg-foreground text-background hover:bg-foreground/90"
          >
            {status === 'loading' ? (
              <>
                <Spinner className="mr-2" />
                Sending...
              </>
            ) : (
              'Send Sign-In Link'
            )}
          </Button>
        </form>
        
        <div className="rule-single my-8" />
        
        <p className="text-center text-sm text-muted-foreground">
          New subscriber?{' '}
          <Link href="/#pricing" className="editorial-link">
            View subscription rates
          </Link>
        </p>
      </div>
    </div>
  )
}

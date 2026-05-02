'use client'

import { useCallback, useState } from 'react'
import { useRouter } from 'next/navigation'
import {
  EmbeddedCheckout,
  EmbeddedCheckoutProvider,
} from '@stripe/react-stripe-js'
import { loadStripe } from '@stripe/stripe-js'

import { startCheckoutSession } from '@/app/actions/stripe'

const stripePromise = loadStripe(process.env.NEXT_PUBLIC_STRIPE_PUBLISHABLE_KEY!)

interface CheckoutProps {
  productId: string
  onComplete?: () => void
}

export default function Checkout({ productId, onComplete }: CheckoutProps) {
  const router = useRouter()
  const [isComplete, setIsComplete] = useState(false)

  const startCheckoutSessionForProduct = useCallback(
    () => startCheckoutSession(productId),
    [productId]
  )

  const handleComplete = useCallback(() => {
    setIsComplete(true)
    if (onComplete) {
      onComplete()
    }
    // Redirect to dashboard after a short delay
    setTimeout(() => {
      router.push('/dashboard')
    }, 2000)
  }, [onComplete, router])

  if (isComplete) {
    return (
      <div className="text-center py-12">
        <h2 className="font-serif text-2xl font-semibold mb-4">
          Welcome to The Morning Tape
        </h2>
        <p className="text-muted-foreground">
          Your subscription is now active. Redirecting to your briefing...
        </p>
      </div>
    )
  }

  return (
    <div id="checkout" className="w-full">
      <EmbeddedCheckoutProvider
        stripe={stripePromise}
        options={{ 
          clientSecret: startCheckoutSessionForProduct,
          onComplete: handleComplete,
        }}
      >
        <EmbeddedCheckout />
      </EmbeddedCheckoutProvider>
    </div>
  )
}

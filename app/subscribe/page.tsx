import { redirect } from 'next/navigation'
import Link from 'next/link'
import { getSession } from '@/lib/auth'
import { PRODUCTS, getProductByPlanType } from '@/lib/products'
import Checkout from '@/components/checkout'
import { Check } from 'lucide-react'

interface SubscribePageProps {
  searchParams: Promise<{ plan?: string }>
}

export default async function SubscribePage({ searchParams }: SubscribePageProps) {
  const session = await getSession()
  
  if (!session) {
    redirect('/login')
  }

  // If user already has subscription, redirect to dashboard
  if (session.subscription) {
    redirect('/dashboard')
  }

  const params = await searchParams
  const planType = params.plan as 'trader' | 'professional' | undefined
  const selectedProduct = planType ? getProductByPlanType(planType) : PRODUCTS[1] // Default to Professional

  if (!selectedProduct) {
    redirect('/subscribe?plan=professional')
  }

  const otherProduct = PRODUCTS.find(p => p.id !== selectedProduct.id)

  return (
    <div className="paper-texture min-h-screen py-12">
      <div className="max-w-4xl mx-auto px-6">
        <div className="text-center mb-8">
          <Link href="/" className="inline-block">
            <h1 className="masthead text-2xl font-serif font-semibold tracking-wider mb-2">
              The Morning Tape
            </h1>
          </Link>
          <div className="rule-double pb-4 mb-4" />
          <p className="section-marker">— Complete Your Subscription —</p>
        </div>

        <div className="grid md:grid-cols-2 gap-8">
          {/* Product Summary */}
          <div className="order-2 md:order-1">
            <div className="border border-border p-6 bg-background">
              <h2 className="font-serif text-2xl font-semibold mb-2">
                {selectedProduct.name}
              </h2>
              <p className="text-muted-foreground text-sm mb-4">
                {selectedProduct.description}
              </p>
              
              <div className="rule-single my-4" />
              
              <p className="font-mono text-3xl font-medium mb-4">
                ${(selectedProduct.priceInCents / 100).toFixed(0)}
                <span className="text-base text-muted-foreground">/month</span>
              </p>
              
              <ul className="space-y-2 mb-6">
                {selectedProduct.features.map((feature, index) => (
                  <li key={index} className="flex items-start gap-3">
                    <Check className="w-4 h-4 text-accent shrink-0 mt-1" />
                    <span className="text-sm">{feature}</span>
                  </li>
                ))}
              </ul>

              {otherProduct && (
                <>
                  <div className="rule-single my-4" />
                  <p className="text-sm text-muted-foreground">
                    Looking for something different?{' '}
                    <Link 
                      href={`/subscribe?plan=${otherProduct.planType}`}
                      className="editorial-link"
                    >
                      Switch to {otherProduct.name}
                    </Link>
                  </p>
                </>
              )}
            </div>
          </div>

          {/* Checkout Form */}
          <div className="order-1 md:order-2">
            <div className="border border-border p-6 bg-background">
              <h3 className="font-serif text-lg font-medium mb-4">
                Payment Details
              </h3>
              <Checkout productId={selectedProduct.id} />
            </div>
          </div>
        </div>

        <div className="rule-single mt-12 mb-8" />
        
        <p className="text-center text-sm text-muted-foreground">
          You can cancel your subscription at any time from your account settings.
        </p>
      </div>
    </div>
  )
}

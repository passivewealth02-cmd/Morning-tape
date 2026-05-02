import Link from 'next/link'
import { Button } from '@/components/ui/button'
import { PRODUCTS } from '@/lib/products'
import { Check } from 'lucide-react'

export function Pricing() {
  return (
    <section className="py-16 md:py-24">
      <div className="max-w-5xl mx-auto px-6">
        <div className="text-center mb-12">
          <p className="section-marker mb-4">— Section III —</p>
          <h2 className="font-serif text-3xl md:text-4xl font-semibold mb-4">
            Subscription Rates
          </h2>
          <p className="text-muted-foreground max-w-xl mx-auto">
            Select the edition that suits your requirements. Both include daily delivery before market open.
          </p>
          <div className="rule-single max-w-xs mx-auto mt-6" />
        </div>
        
        <div className="grid md:grid-cols-2 gap-8 mt-12">
          {PRODUCTS.map((product, index) => (
            <article 
              key={product.id} 
              className={`relative border ${index === 1 ? 'border-accent border-2' : 'border-border'} bg-background p-8`}
            >
              {index === 1 && (
                <div className="absolute -top-3 left-1/2 -translate-x-1/2 bg-accent text-accent-foreground px-4 py-1 text-xs tracking-widest uppercase font-serif">
                  Recommended
                </div>
              )}
              
              <div className="text-center mb-6">
                <h3 className="font-serif text-2xl font-semibold mb-2">
                  {product.name}
                </h3>
                <p className="text-muted-foreground text-sm mb-4">
                  {product.description}
                </p>
                <div className="rule-single mb-4" />
                <p className="font-mono text-4xl font-medium">
                  ${(product.priceInCents / 100).toFixed(0)}
                  <span className="text-lg text-muted-foreground">/month</span>
                </p>
              </div>
              
              <ul className="space-y-3 mb-8">
                {product.features.map((feature, featureIndex) => (
                  <li key={featureIndex} className="flex items-start gap-3">
                    <Check className="w-5 h-5 text-accent shrink-0 mt-0.5" />
                    <span className="text-sm">{feature}</span>
                  </li>
                ))}
              </ul>
              
              <Link href={`/login?plan=${product.planType}`} className="block">
                <Button 
                  className={`w-full font-serif tracking-wide ${
                    index === 1 
                      ? 'bg-accent text-accent-foreground hover:bg-accent/90' 
                      : 'bg-foreground text-background hover:bg-foreground/90'
                  }`}
                >
                  Subscribe to {product.name}
                </Button>
              </Link>
            </article>
          ))}
        </div>
      </div>
    </section>
  )
}

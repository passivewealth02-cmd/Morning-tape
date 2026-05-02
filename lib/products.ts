export interface Product {
  id: string
  name: string
  description: string
  priceInCents: number
  planType: 'trader' | 'professional'
  features: string[]
}

export const PRODUCTS: Product[] = [
  {
    id: 'the-trader',
    name: 'The Trader',
    description: 'Essential daily market intelligence for the active investor',
    priceInCents: 2900, // $29/month
    planType: 'trader',
    features: [
      'Daily market overview',
      'Top gainers and losers with sparkline charts',
      'AI-generated market commentary',
      'Email delivery before market open',
    ],
  },
  {
    id: 'the-professional',
    name: 'The Professional',
    description: 'Comprehensive analysis for the discerning market participant',
    priceInCents: 4900, // $49/month
    planType: 'professional',
    features: [
      'Everything in The Trader',
      'Sector performance breakdown',
      'Economic calendar with importance ratings',
      'Extended AI analysis and insights',
      'Priority email delivery',
    ],
  },
]

export function getProductById(id: string): Product | undefined {
  return PRODUCTS.find((p) => p.id === id)
}

export function getProductByPlanType(planType: 'trader' | 'professional'): Product | undefined {
  return PRODUCTS.find((p) => p.planType === planType)
}

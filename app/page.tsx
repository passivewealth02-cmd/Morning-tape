import { Masthead } from '@/components/landing/masthead'
import { Hero } from '@/components/landing/hero'
import { Features } from '@/components/landing/features'
import { Pricing } from '@/components/landing/pricing'
import { Footer } from '@/components/landing/footer'

export default function Home() {
  return (
    <div className="paper-texture min-h-screen">
      <main className="max-w-7xl mx-auto">
        <Masthead />
        <Hero />
        <Features />
        <Pricing />
      </main>
      <Footer />
    </div>
  )
}

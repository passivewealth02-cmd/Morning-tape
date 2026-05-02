import Link from 'next/link'
import { Button } from '@/components/ui/button'

export function Hero() {
  return (
    <section className="max-w-4xl mx-auto px-6 py-12 md:py-16">
      <div className="text-center mb-12">
        <p className="section-marker mb-6">— The Daily Brief —</p>
      </div>
      
      <article className="prose prose-lg max-w-none">
        <p className="drop-cap text-xl md:text-2xl leading-relaxed mb-8 font-sans">
          Before the opening bell, before the cable news cycle, before the noise—there is The Morning Tape. 
          A single, authoritative page of AI-curated market intelligence delivered to your inbox each trading day.
        </p>
        
        <div className="rule-single my-8" />
        
        <div className="grid md:grid-cols-2 gap-8 text-base leading-relaxed">
          <div>
            <p className="mb-4">
              Our artificial intelligence synthesizes overnight market movements, pre-market indicators, 
              economic releases, and global sentiment into a concise morning briefing.
            </p>
            <p>
              No advertisements. No sponsored content. No algorithmic feed. Just the essential 
              intelligence you need to start your trading day with clarity.
            </p>
          </div>
          <div>
            <p className="mb-4">
              Each edition features top movers with historical sparklines, sector rotation analysis, 
              and an AI commentary section that cuts through the market chatter.
            </p>
            <p>
              Join the discerning investors who begin their day not with a deluge of data, 
              but with a single, well-crafted page of market wisdom.
            </p>
          </div>
        </div>
      </article>
      
      <div className="rule-single my-12" />
      
      <div className="text-center">
        <Link href="/login">
          <Button 
            size="lg" 
            className="font-serif tracking-wide text-lg px-12 py-6 bg-foreground text-background hover:bg-foreground/90"
          >
            Begin Your Subscription
          </Button>
        </Link>
        <p className="mt-4 text-sm text-muted-foreground">
          No commitment required. Cancel anytime.
        </p>
      </div>
    </section>
  )
}

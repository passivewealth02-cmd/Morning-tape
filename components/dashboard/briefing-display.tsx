'use client'

import useSWR from 'swr'
import type { BriefingContent } from '@/lib/db'
import { Sparkline } from './sparkline'
import { SectionDivider } from './section-divider'
import { StarRating } from './star-rating'
import { Spinner } from '@/components/ui/spinner'

const fetcher = (url: string) => fetch(url).then((res) => res.json())

interface BriefingDisplayProps {
  plan: 'trader' | 'professional'
}

export function BriefingDisplay({ plan }: BriefingDisplayProps) {
  const { data, error, isLoading } = useSWR<{ briefing: BriefingContent; cached: boolean }>(
    '/api/generate-brief',
    fetcher,
    {
      revalidateOnFocus: false,
      revalidateOnReconnect: false,
    }
  )

  if (isLoading) {
    return (
      <div className="flex flex-col items-center justify-center py-24">
        <Spinner className="w-8 h-8 mb-4" />
        <p className="text-muted-foreground font-serif">
          Preparing your briefing...
        </p>
      </div>
    )
  }

  if (error || !data?.briefing) {
    return (
      <div className="text-center py-24">
        <p className="text-accent font-serif text-lg">
          Unable to load today&apos;s briefing.
        </p>
        <p className="text-muted-foreground mt-2">
          Please refresh the page to try again.
        </p>
      </div>
    )
  }

  const { briefing } = data

  return (
    <article className="max-w-4xl mx-auto">
      {/* Market Overview */}
      <SectionDivider section="— Section I —" />
      <section>
        <h2 className="font-serif text-2xl font-semibold mb-6 text-center">
          Market Overview
        </h2>
        <div className="drop-cap text-lg leading-relaxed whitespace-pre-line">
          {briefing.marketOverview}
        </div>
      </section>

      {/* Top Movers */}
      <SectionDivider section="— Section II —" />
      <section>
        <h2 className="font-serif text-2xl font-semibold mb-6 text-center">
          Top Movers
        </h2>
        
        <div className="grid md:grid-cols-2 gap-8">
          {/* Gainers */}
          <div>
            <h3 className="font-serif text-lg font-medium mb-4 text-center">
              Gainers
            </h3>
            <div className="space-y-3">
              {briefing.topMovers.gainers.map((stock) => (
                <div 
                  key={stock.ticker}
                  className="flex items-center justify-between p-3 border border-border"
                >
                  <div className="flex items-center gap-3">
                    <span className="ticker text-sm">{stock.ticker}</span>
                    <Sparkline data={stock.sparklineData} trend="up" />
                  </div>
                  <div className="text-right">
                    <p className="font-mono text-sm">${stock.price.toFixed(2)}</p>
                    <p className="text-sm text-green-700">{stock.change}</p>
                  </div>
                </div>
              ))}
            </div>
          </div>

          {/* Losers */}
          <div>
            <h3 className="font-serif text-lg font-medium mb-4 text-center">
              Losers
            </h3>
            <div className="space-y-3">
              {briefing.topMovers.losers.map((stock) => (
                <div 
                  key={stock.ticker}
                  className="flex items-center justify-between p-3 border border-border"
                >
                  <div className="flex items-center gap-3">
                    <span className="ticker text-sm">{stock.ticker}</span>
                    <Sparkline data={stock.sparklineData} trend="down" />
                  </div>
                  <div className="text-right">
                    <p className="font-mono text-sm">${stock.price.toFixed(2)}</p>
                    <p className="text-sm text-accent">{stock.change}</p>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      </section>

      {/* Sector Performance (Professional only) */}
      {plan === 'professional' && briefing.sectorPerformance && (
        <>
          <SectionDivider section="— Section III —" />
          <section>
            <h2 className="font-serif text-2xl font-semibold mb-6 text-center">
              Sector Performance
            </h2>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
              {briefing.sectorPerformance.map((sector) => (
                <div 
                  key={sector.sector}
                  className="p-4 border border-border text-center"
                >
                  <p className="text-sm text-muted-foreground mb-1">{sector.sector}</p>
                  <p className={`font-mono font-medium ${
                    sector.trend === 'up' ? 'text-green-700' : 
                    sector.trend === 'down' ? 'text-accent' : 
                    'text-muted-foreground'
                  }`}>
                    {sector.change}
                  </p>
                </div>
              ))}
            </div>
          </section>
        </>
      )}

      {/* Economic Calendar (Professional only) */}
      {plan === 'professional' && briefing.economicCalendar && (
        <>
          <SectionDivider section="— Section IV —" />
          <section>
            <h2 className="font-serif text-2xl font-semibold mb-6 text-center">
              Economic Calendar
            </h2>
            <div className="space-y-3">
              {briefing.economicCalendar.map((event, index) => (
                <div 
                  key={index}
                  className="flex items-center justify-between p-4 border border-border"
                >
                  <div className="flex items-center gap-4">
                    <StarRating rating={event.importance} />
                    <span className="font-medium">{event.event}</span>
                  </div>
                  <span className="font-mono text-sm text-muted-foreground">
                    {event.time}
                  </span>
                </div>
              ))}
            </div>
          </section>
        </>
      )}

      {/* AI Commentary */}
      <SectionDivider section="— Commentary —" />
      <section className="mb-12">
        <h2 className="font-serif text-2xl font-semibold mb-6 text-center">
          Market Analysis
        </h2>
        <div className="drop-cap text-lg leading-relaxed whitespace-pre-line">
          {briefing.aiCommentary}
        </div>
      </section>
    </article>
  )
}

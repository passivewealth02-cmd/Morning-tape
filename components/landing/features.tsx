export function Features() {
  const features = [
    {
      title: 'Market Overview',
      description: 'A synthesis of overnight developments, futures positioning, and pre-market sentiment distilled into essential paragraphs.',
    },
    {
      title: 'Top Movers',
      description: 'The day&apos;s most significant price movements, each accompanied by a seven-day sparkline chart for context.',
    },
    {
      title: 'Sector Analysis',
      description: 'Rotation patterns and relative strength across market sectors, identifying where institutional capital flows.',
    },
    {
      title: 'Economic Calendar',
      description: 'Upcoming releases and events rated by market impact, so you know what may move prices today.',
    },
    {
      title: 'AI Commentary',
      description: 'Machine intelligence parsing through noise to surface the narratives that matter to your portfolio.',
    },
    {
      title: 'Inbox Delivery',
      description: 'Arrives before market open. Read on any device. Archive for reference. No app required.',
    },
  ]

  return (
    <section className="py-16 md:py-24 bg-card">
      <div className="max-w-6xl mx-auto px-6">
        <div className="text-center mb-12">
          <p className="section-marker mb-4">— Section II —</p>
          <h2 className="font-serif text-3xl md:text-4xl font-semibold mb-4">
            What You Receive
          </h2>
          <div className="rule-single max-w-xs mx-auto" />
        </div>
        
        <div className="grid md:grid-cols-2 lg:grid-cols-3 gap-8 mt-12">
          {features.map((feature, index) => (
            <article key={index} className="group">
              <div className="border-t border-border pt-6">
                <h3 className="font-serif text-xl font-medium mb-3">
                  {feature.title}
                </h3>
                <p className="text-muted-foreground leading-relaxed">
                  {feature.description}
                </p>
              </div>
            </article>
          ))}
        </div>
      </div>
    </section>
  )
}

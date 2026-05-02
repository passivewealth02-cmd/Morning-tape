export function Masthead() {
  // Calculate volume and issue number based on date
  const today = new Date()
  const startDate = new Date('2024-01-01')
  const daysSinceStart = Math.floor((today.getTime() - startDate.getTime()) / (1000 * 60 * 60 * 24))
  const issueNumber = daysSinceStart + 1
  const volume = Math.floor(daysSinceStart / 365) + 3 // Starting at Vol. III

  return (
    <header className="text-center py-8 md:py-12">
      <div className="rule-double pb-6 mb-6">
        <p className="text-xs tracking-[0.3em] text-muted-foreground mb-2 uppercase">
          Est. 2024 — Market Intelligence
        </p>
      </div>
      
      <h1 className="masthead text-4xl md:text-6xl lg:text-7xl font-serif font-semibold tracking-wider mb-4">
        The Morning Tape
      </h1>
      
      <p className="text-sm md:text-base tracking-[0.2em] text-muted-foreground font-serif">
        Vol. {toRoman(volume)} · No. {issueNumber}
      </p>
      
      <div className="rule-double mt-6 pt-1" />
    </header>
  )
}

function toRoman(num: number): string {
  const romanNumerals: [number, string][] = [
    [10, 'X'],
    [9, 'IX'],
    [5, 'V'],
    [4, 'IV'],
    [1, 'I'],
  ]
  
  let result = ''
  let remaining = num
  
  for (const [value, symbol] of romanNumerals) {
    while (remaining >= value) {
      result += symbol
      remaining -= value
    }
  }
  
  return result
}

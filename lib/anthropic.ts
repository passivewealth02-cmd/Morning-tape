import 'server-only'
import Anthropic from '@anthropic-ai/sdk'
import type { BriefingContent } from './db'

const anthropic = new Anthropic({
  apiKey: process.env.ANTHROPIC_API_KEY,
})

export async function generateBriefing(plan: 'trader' | 'professional'): Promise<BriefingContent> {
  const today = new Date()
  const dateStr = today.toLocaleDateString('en-US', { 
    weekday: 'long', 
    year: 'numeric', 
    month: 'long', 
    day: 'numeric' 
  })

  const professionalSections = plan === 'professional' ? `
  - "sectorPerformance": Array of sector data with "sector" (string), "change" (string like "+1.2%"), "trend" ("up" | "down" | "flat")
  - "economicCalendar": Array of upcoming events with "event" (string), "time" (string), "importance" (1, 2, or 3 for star rating)
  ` : ''

  const prompt = `You are a financial analyst writing the daily market briefing for The Morning Tape, a premium editorial-style market intelligence publication. Today's date is ${dateStr}.

Generate a market briefing in JSON format. Write in a professional, authoritative tone befitting a publication like the Financial Times or a hedge fund morning note. Be concise but insightful.

The JSON structure must be exactly:
{
  "marketOverview": "A 2-3 paragraph overview of current market conditions, overnight developments, and the trading day ahead. Write with authority and insight.",
  "topMovers": {
    "gainers": [
      {
        "ticker": "SYMBOL",
        "name": "Company Name",
        "change": "+X.X%",
        "price": 123.45,
        "sparklineData": [array of 7 numbers representing last 7 days normalized price movement, values between 0-100]
      }
    ],
    "losers": [same structure as gainers]
  },
  ${professionalSections}
  "aiCommentary": "A thoughtful 2-3 paragraph analysis that cuts through market noise to identify the key narrative or theme of the day. This should feel like insight from a seasoned market strategist."
}

Include 3-4 gainers and 3-4 losers with realistic current market data. ${plan === 'professional' ? 'Include 6-8 sectors and 4-6 economic calendar events.' : ''}

The sparklineData should be 7 numbers showing realistic price movement patterns - upward trending for gainers, downward for losers.

Return ONLY valid JSON, no additional text or markdown.`

  const message = await anthropic.messages.create({
    model: 'claude-sonnet-4-20250514',
    max_tokens: 2048,
    messages: [
      {
        role: 'user',
        content: prompt,
      },
    ],
  })

  const content = message.content[0]
  if (content.type !== 'text') {
    throw new Error('Unexpected response type from Anthropic')
  }

  try {
    const briefing = JSON.parse(content.text) as BriefingContent
    return briefing
  } catch (error) {
    console.error('Failed to parse briefing JSON:', content.text)
    throw new Error('Failed to parse briefing response')
  }
}

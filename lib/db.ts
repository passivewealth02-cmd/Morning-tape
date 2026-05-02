import { neon } from '@neondatabase/serverless'

export const sql = neon(process.env.DATABASE_URL!)

export type User = {
  id: string
  email: string
  created_at: string
  updated_at: string
}

export type Subscription = {
  id: string
  user_id: string
  stripe_customer_id: string | null
  stripe_subscription_id: string | null
  plan: 'trader' | 'professional'
  status: 'active' | 'canceled' | 'past_due' | 'trialing'
  current_period_end: string | null
  created_at: string
  updated_at: string
}

export type Briefing = {
  id: string
  date: string
  plan: 'trader' | 'professional'
  content: BriefingContent
  created_at: string
}

export type BriefingContent = {
  marketOverview: string
  topMovers: {
    gainers: TickerData[]
    losers: TickerData[]
  }
  sectorPerformance?: SectorData[]
  economicCalendar?: CalendarEvent[]
  aiCommentary: string
}

export type TickerData = {
  ticker: string
  name: string
  change: string
  price: number
  sparklineData: number[]
}

export type SectorData = {
  sector: string
  change: string
  trend: 'up' | 'down' | 'flat'
}

export type CalendarEvent = {
  event: string
  time: string
  importance: 1 | 2 | 3
}

export type Session = {
  id: string
  user_id: string
  token: string
  expires_at: string
  created_at: string
}

import { NextRequest, NextResponse } from 'next/server'
import { getSession } from '@/lib/auth'
import { sql, type BriefingContent } from '@/lib/db'
import { generateBriefing } from '@/lib/anthropic'

export async function GET(request: NextRequest) {
  try {
    const session = await getSession()
    
    if (!session || !session.subscription) {
      return NextResponse.json(
        { error: 'Unauthorized' },
        { status: 401 }
      )
    }

    const plan = session.subscription.plan
    const today = new Date().toISOString().split('T')[0] // YYYY-MM-DD

    // Check for cached briefing
    const cachedBriefing = await sql`
      SELECT content FROM briefings
      WHERE date = ${today}
        AND plan = ${plan}
      LIMIT 1
    `

    if (cachedBriefing.length > 0) {
      return NextResponse.json({
        briefing: cachedBriefing[0].content as BriefingContent,
        cached: true,
      })
    }

    // Generate new briefing
    const briefing = await generateBriefing(plan)

    // Cache the briefing
    await sql`
      INSERT INTO briefings (date, plan, content)
      VALUES (${today}, ${plan}, ${JSON.stringify(briefing)})
      ON CONFLICT (date, plan) DO UPDATE SET content = ${JSON.stringify(briefing)}
    `

    return NextResponse.json({
      briefing,
      cached: false,
    })
  } catch (error) {
    console.error('Error generating briefing:', error)
    return NextResponse.json(
      { error: 'Failed to generate briefing' },
      { status: 500 }
    )
  }
}

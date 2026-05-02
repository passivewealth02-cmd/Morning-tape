import { NextResponse } from 'next/server'
import { sql } from '@/lib/db'
import { generateBriefing } from '@/lib/anthropic'
import { Resend } from 'resend'

const resend = new Resend(process.env.RESEND_API_KEY)

export async function GET(request: Request) {
  // Verify cron secret (Vercel sends this automatically)
  const authHeader = request.headers.get('authorization')
  if (authHeader !== `Bearer ${process.env.CRON_SECRET}`) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
  }

  try {
    const today = new Date().toISOString().split('T')[0]
    const dateStr = new Date().toLocaleDateString('en-US', {
      weekday: 'long', month: 'long', day: 'numeric', year: 'numeric'
    })

    // Generate briefs for both tiers
    const [traderBrief, proBrief] = await Promise.all([
      generateBriefing('trader'),
      generateBriefing('professional'),
    ])

    // Cache briefs
    await sql`
      INSERT INTO briefings (date, plan, content)
      VALUES (${today}, 'trader', ${JSON.stringify(traderBrief)})
      ON CONFLICT (date, plan) DO UPDATE SET content = ${JSON.stringify(traderBrief)}
    `
    await sql`
      INSERT INTO briefings (date, plan, content)
      VALUES (${today}, 'professional', ${JSON.stringify(proBrief)})
      ON CONFLICT (date, plan) DO UPDATE SET content = ${JSON.stringify(proBrief)}
    `

    // Get all active subscribers
    const subscribers = await sql`
      SELECT u.email, s.plan
      FROM users u
      JOIN subscriptions s ON u.id = s.user_id
      WHERE s.status = 'active'
    `

    // Send emails
    let sent = 0
    for (const sub of subscribers) {
      const brief = sub.plan === 'professional' ? proBrief : traderBrief

      try {
        await resend.emails.send({
          from: 'The Morning Tape <noreply@themorningtape.com>',
          to: sub.email,
          subject: `The Morning Tape — ${dateStr}`,
          html: buildEmailHtml(brief, sub.plan, dateStr),
        })
        sent++
      } catch (err) {
        console.error(`Failed to send to ${sub.email}:`, err)
      }
    }

    return NextResponse.json({ success: true, sent, total: subscribers.length })
  } catch (error) {
    console.error('Cron error:', error)
    return NextResponse.json({ error: 'Cron failed' }, { status: 500 })
  }
}

function buildEmailHtml(brief: any, plan: string, dateStr: string): string {
  const gainers = brief.topMovers?.gainers?.map((t: any) =>
    `<tr><td style="font-family:monospace;font-weight:600;padding:8px 16px 8px 0;">${t.ticker}</td><td style="color:#228B22;font-family:monospace;">${t.change}</td><td style="padding-left:16px;">${t.name}</td></tr>`
  ).join('') || ''

  const losers = brief.topMovers?.losers?.map((t: any) =>
    `<tr><td style="font-family:monospace;font-weight:600;padding:8px 16px 8px 0;">${t.ticker}</td><td style="color:#8B0000;font-family:monospace;">${t.change}</td><td style="padding-left:16px;">${t.name}</td></tr>`
  ).join('') || ''

  const calendar = plan === 'professional' && brief.economicCalendar
    ? brief.economicCalendar.map((e: any) =>
        `<tr><td style="font-family:monospace;font-size:13px;padding:6px 12px 6px 0;color:#5C5549;">${e.time}</td><td style="padding:6px 0;">${e.event}</td><td style="text-align:right;color:#8B0000;">${'★'.repeat(e.importance)}</td></tr>`
      ).join('')
    : ''

  return `
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="font-family:Georgia,serif;background:#F4EDE0;padding:40px 20px;margin:0;">
<div style="max-width:640px;margin:0 auto;">
  <div style="text-align:center;padding-bottom:20px;border-bottom:3px double #1A1612;">
    <p style="font-size:11px;letter-spacing:0.3em;text-transform:uppercase;color:#5C5549;margin:0 0 8px;">⁂ Established for the Discerning Trader ⁂</p>
    <h1 style="font-family:Georgia,serif;font-size:36px;letter-spacing:0.1em;text-transform:uppercase;color:#1A1612;margin:0;">The Morning Tape</h1>
    <p style="font-size:13px;color:#5C5549;margin:8px 0 0;letter-spacing:0.15em;">${dateStr}</p>
  </div>

  <div style="padding:28px 0;border-bottom:1px solid #C9BFAB;">
    <p style="font-size:11px;letter-spacing:0.2em;text-transform:uppercase;color:#8B0000;margin:0 0 12px;">⸻ Market Overview ⸻</p>
    <p style="font-size:17px;line-height:1.65;color:#1A1612;margin:0;">${brief.marketOverview}</p>
  </div>

  <div style="padding:28px 0;border-bottom:1px solid #C9BFAB;">
    <p style="font-size:11px;letter-spacing:0.2em;text-transform:uppercase;color:#5C5549;margin:0 0 16px;">⁘ Top Gainers ⁘</p>
    <table style="width:100%;border-collapse:collapse;">${gainers}</table>
    <p style="font-size:11px;letter-spacing:0.2em;text-transform:uppercase;color:#5C5549;margin:24px 0 16px;">⁘ Top Losers ⁘</p>
    <table style="width:100%;border-collapse:collapse;">${losers}</table>
  </div>

  ${calendar ? `
  <div style="padding:28px 0;border-bottom:1px solid #C9BFAB;">
    <p style="font-size:11px;letter-spacing:0.2em;text-transform:uppercase;color:#5C5549;margin:0 0 16px;">⁘ Economic Calendar ⁘</p>
    <table style="width:100%;border-collapse:collapse;">${calendar}</table>
  </div>
  ` : ''}

  <div style="padding:28px 0;border-bottom:3px double #1A1612;">
    <p style="font-size:11px;letter-spacing:0.2em;text-transform:uppercase;color:#8B0000;margin:0 0 12px;">⸻ The Editor's Take ⸻</p>
    <p style="font-size:17px;line-height:1.65;color:#1A1612;margin:0;font-style:italic;">${brief.aiCommentary}</p>
  </div>

  <div style="text-align:center;padding:24px 0;">
    <p style="font-size:11px;letter-spacing:0.2em;text-transform:uppercase;color:#5C5549;margin:0;">⁂ End of Edition ⁂</p>
    <p style="font-size:12px;color:#5C5549;margin:12px 0 0;">
      <a href="https://themorningtape.com/dashboard" style="color:#8B0000;">View on Web</a> · 
      <a href="https://themorningtape.com" style="color:#8B0000;">Manage Subscription</a>
    </p>
  </div>
</div>
</body>
</html>`
}

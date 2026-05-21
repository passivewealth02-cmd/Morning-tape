import Link from 'next/link'
import { ArrowRight, Check, X, Minus } from 'lucide-react'
import { SiteHeader } from '@/components/marketing/site-header'
import { SiteFooter } from '@/components/marketing/site-footer'
import { pageMetadata } from '@/lib/seo'

export const metadata = pageMetadata({
  title: 'Maintena vs. Alternatives — Maintenance Software Comparison',
  description:
    'Compare Maintena to Buildium, AppFolio, PropertyMeld, and spreadsheets. See which property maintenance software gives you the most for your money.',
  path: '/compare',
  keywords: [
    'property maintenance software comparison',
    'maintena vs buildium',
    'maintena vs appfolio',
    'best maintenance management software',
    'property maintenance software alternatives',
  ],
})

const competitors = [
  { id: 'maintena', name: 'Maintena', highlight: true, price: 'from $99/mo', note: '14-day free trial' },
  { id: 'spreadsheet', name: 'Spreadsheets + Email', highlight: false, price: '$0', note: 'Your time isn\'t free' },
  { id: 'buildium', name: 'Buildium', highlight: false, price: 'from $55/mo', note: 'Full PM suite — charged per unit' },
  { id: 'appfolio', name: 'AppFolio', highlight: false, price: '$1.40–$3/unit/mo', note: '$280/mo minimum' },
  { id: 'propertymeld', name: 'PropertyMeld', highlight: false, price: 'from $160/mo', note: 'Maintenance-only, no AI' },
]

type Mark = 'yes' | 'no' | 'partial'

interface Row {
  category: string
  feature: string
  detail: string
  values: Record<string, Mark>
}

const rows: Row[] = [
  // Intake
  {
    category: 'Intake',
    feature: 'Tenant email submission',
    detail: 'Tenants submit by email, no account or app needed',
    values: { maintena: 'yes', spreadsheet: 'partial', buildium: 'partial', appfolio: 'yes', propertymeld: 'yes' },
  },
  {
    category: 'Intake',
    feature: 'Web form for tenants',
    detail: 'Embeddable or hosted form tenants can fill out',
    values: { maintena: 'yes', spreadsheet: 'no', buildium: 'yes', appfolio: 'yes', propertymeld: 'yes' },
  },
  {
    category: 'Intake',
    feature: 'No tenant app required',
    detail: 'Tenants never need to create an account or download anything',
    values: { maintena: 'yes', spreadsheet: 'yes', buildium: 'no', appfolio: 'no', propertymeld: 'partial' },
  },
  {
    category: 'Intake',
    feature: 'Photo & file attachments',
    detail: 'Tenants and managers can attach photos, PDFs to tickets',
    values: { maintena: 'yes', spreadsheet: 'no', buildium: 'yes', appfolio: 'yes', propertymeld: 'yes' },
  },
  {
    category: 'Intake',
    feature: 'Private, secure file storage',
    detail: 'Attachments never publicly accessible — served via authenticated proxy',
    values: { maintena: 'yes', spreadsheet: 'no', buildium: 'partial', appfolio: 'partial', propertymeld: 'no' },
  },
  // AI
  {
    category: 'AI & Automation',
    feature: 'AI trade classification',
    detail: 'Automatically identifies plumbing, electrical, HVAC, etc.',
    values: { maintena: 'yes', spreadsheet: 'no', buildium: 'no', appfolio: 'no', propertymeld: 'no' },
  },
  {
    category: 'AI & Automation',
    feature: 'AI urgency scoring',
    detail: 'Flags emergencies and high-risk issues instantly',
    values: { maintena: 'yes', spreadsheet: 'no', buildium: 'no', appfolio: 'no', propertymeld: 'no' },
  },
  {
    category: 'AI & Automation',
    feature: 'AI-written ticket summaries',
    detail: 'Plain-English summary auto-generated from raw tenant text',
    values: { maintena: 'yes', spreadsheet: 'no', buildium: 'no', appfolio: 'no', propertymeld: 'no' },
  },
  {
    category: 'AI & Automation',
    feature: 'AI vendor recommendations',
    detail: 'Suggests best vendor by trade, location, availability, and performance',
    values: { maintena: 'yes', spreadsheet: 'no', buildium: 'no', appfolio: 'partial', propertymeld: 'no' },
  },
  {
    category: 'AI & Automation',
    feature: 'Automatic escalation detection',
    detail: 'Flags tickets likely to escalate before they become emergencies',
    values: { maintena: 'yes', spreadsheet: 'no', buildium: 'no', appfolio: 'no', propertymeld: 'no' },
  },
  // Vendors
  {
    category: 'Vendors',
    feature: 'Bring your own vendors',
    detail: 'Add and manage your existing vendor network',
    values: { maintena: 'yes', spreadsheet: 'yes', buildium: 'yes', appfolio: 'yes', propertymeld: 'yes' },
  },
  {
    category: 'Vendors',
    feature: 'Vendor performance tracking',
    detail: 'Ratings and completion data tracked per vendor',
    values: { maintena: 'yes', spreadsheet: 'no', buildium: 'partial', appfolio: 'yes', propertymeld: 'yes' },
  },
  {
    category: 'Vendors',
    feature: 'One-click vendor notification',
    detail: 'Notify the assigned vendor by email in one click',
    values: { maintena: 'yes', spreadsheet: 'no', buildium: 'yes', appfolio: 'yes', propertymeld: 'yes' },
  },
  {
    category: 'Vendors',
    feature: 'Insurance status tracking',
    detail: 'Track vendor insurance verification status',
    values: { maintena: 'yes', spreadsheet: 'no', buildium: 'partial', appfolio: 'yes', propertymeld: 'partial' },
  },
  // Tracking
  {
    category: 'Tracking & Reporting',
    feature: 'Kanban board view',
    detail: 'Drag-and-drop status board across New, Assigned, In Progress, Done',
    values: { maintena: 'yes', spreadsheet: 'no', buildium: 'no', appfolio: 'no', propertymeld: 'yes' },
  },
  {
    category: 'Tracking & Reporting',
    feature: 'SLA timers & alerts',
    detail: 'Configurable response-time targets with alerts before breach',
    values: { maintena: 'yes', spreadsheet: 'no', buildium: 'partial', appfolio: 'yes', propertymeld: 'yes' },
  },
  {
    category: 'Tracking & Reporting',
    feature: 'Full audit timeline',
    detail: 'Every status change, message, and action logged with timestamps',
    values: { maintena: 'yes', spreadsheet: 'no', buildium: 'partial', appfolio: 'yes', propertymeld: 'partial' },
  },
  {
    category: 'Tracking & Reporting',
    feature: 'Automated tenant updates',
    detail: 'Status emails sent to tenants at each stage — no manual follow-up',
    values: { maintena: 'yes', spreadsheet: 'no', buildium: 'partial', appfolio: 'yes', propertymeld: 'yes' },
  },
  // Value
  {
    category: 'Value & Setup',
    feature: 'Setup in under 5 minutes',
    detail: 'Add properties, connect email, and start routing the same day',
    values: { maintena: 'yes', spreadsheet: 'yes', buildium: 'no', appfolio: 'no', propertymeld: 'partial' },
  },
  {
    category: 'Value & Setup',
    feature: 'No long-term contract',
    detail: 'Cancel any time — month-to-month billing',
    values: { maintena: 'yes', spreadsheet: 'yes', buildium: 'partial', appfolio: 'no', propertymeld: 'partial' },
  },
  {
    category: 'Value & Setup',
    feature: 'Free trial',
    detail: 'Try the full product before paying',
    values: { maintena: 'yes', spreadsheet: 'yes', buildium: 'yes', appfolio: 'no', propertymeld: 'partial' },
  },
  {
    category: 'Value & Setup',
    feature: 'Maintenance-focused (not bloated)',
    detail: 'Purpose-built for maintenance coordination — not buried in an accounting suite',
    values: { maintena: 'yes', spreadsheet: 'yes', buildium: 'no', appfolio: 'no', propertymeld: 'yes' },
  },
]

function Mark({ value }: { value: Mark }) {
  if (value === 'yes') return <Check className="w-5 h-5 text-green-500 mx-auto" aria-label="Yes" />
  if (value === 'no') return <X className="w-5 h-5 text-gray-300 mx-auto" aria-label="No" />
  return <Minus className="w-5 h-5 text-yellow-400 mx-auto" aria-label="Partial" />
}

const categories = [...new Set(rows.map(r => r.category))]

const painPoints = [
  {
    from: 'Requests lost in email',
    to: 'Every request becomes a structured ticket automatically',
  },
  {
    from: 'Manual trade triage takes minutes per ticket',
    to: 'AI classifies trade and urgency in seconds',
  },
  {
    from: 'Calling around to find an available vendor',
    to: 'AI ranks your vendors and you assign in one click',
  },
  {
    from: 'Tenants calling "any update?" constantly',
    to: 'Automatic status emails at every stage',
  },
  {
    from: 'No visibility into who did what, when',
    to: 'Full audit timeline — every action logged with timestamps',
  },
  {
    from: 'Tenant photos emailed separately, never attached to the ticket',
    to: 'Private, secure photo attachments directly on the ticket',
  },
]

export default function ComparePage() {
  return (
    <div className="min-h-screen bg-white">
      <SiteHeader />

      <main>
        {/* Hero */}
        <section className="max-w-5xl mx-auto px-6 pt-16 sm:pt-24 pb-12 text-center">
          <div className="inline-flex items-center gap-2 text-xs font-medium bg-indigo-50 text-indigo-700 px-3 py-1.5 rounded-full mb-6">
            Honest comparison
          </div>
          <h1 className="text-3xl sm:text-5xl font-semibold tracking-tight text-gray-900 mb-5">
            Maintena vs. the alternatives
          </h1>
          <p className="text-base sm:text-lg text-gray-500 max-w-2xl mx-auto mb-8">
            Most property maintenance software was built as an afterthought inside a bigger accounting suite. Maintena is built specifically for maintenance coordination — and adds AI where it actually saves you time.
          </p>
          <Link
            href="/signup"
            className="inline-flex items-center gap-2 bg-indigo-600 text-white px-7 py-3 rounded-lg font-medium hover:bg-indigo-700 transition-colors"
          >
            Start free trial — 14 days free
            <ArrowRight className="w-4 h-4" />
          </Link>
        </section>

        {/* Before / after */}
        <section className="bg-gray-50 border-y border-gray-100 py-16 sm:py-20">
          <div className="max-w-4xl mx-auto px-6">
            <h2 className="text-xl sm:text-2xl font-semibold text-gray-900 text-center mb-10">
              What life looks like before and after Maintena
            </h2>
            <div className="space-y-3">
              {painPoints.map(p => (
                <div key={p.from} className="grid sm:grid-cols-2 gap-0 rounded-xl overflow-hidden border border-gray-200">
                  <div className="bg-white px-5 py-4 flex items-start gap-3">
                    <X className="w-4 h-4 text-red-400 mt-0.5 shrink-0" />
                    <p className="text-sm text-gray-600">{p.from}</p>
                  </div>
                  <div className="bg-indigo-50 px-5 py-4 flex items-start gap-3 border-t sm:border-t-0 sm:border-l border-gray-200">
                    <Check className="w-4 h-4 text-green-500 mt-0.5 shrink-0" />
                    <p className="text-sm text-gray-700 font-medium">{p.to}</p>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </section>

        {/* Comparison table */}
        <section className="py-16 sm:py-24">
          <div className="max-w-6xl mx-auto px-4 sm:px-6">
            <h2 className="text-xl sm:text-2xl font-semibold text-gray-900 text-center mb-3">
              Feature-by-feature comparison
            </h2>
            <p className="text-sm text-gray-500 text-center mb-10">
              <Minus className="w-4 h-4 text-yellow-400 inline mb-0.5" /> = partial support &nbsp;·&nbsp;
              <Check className="w-4 h-4 text-green-500 inline mb-0.5" /> = full support &nbsp;·&nbsp;
              <X className="w-4 h-4 text-gray-300 inline mb-0.5" /> = not available
            </p>

            <div className="overflow-x-auto rounded-xl border border-gray-200">
              <table className="w-full text-sm border-collapse">
                <thead>
                  <tr className="border-b border-gray-200">
                    <th className="text-left px-4 py-4 font-medium text-gray-500 bg-gray-50 w-64">Feature</th>
                    {competitors.map(c => (
                      <th
                        key={c.id}
                        className={`px-3 py-4 text-center font-semibold text-sm ${c.highlight ? 'bg-indigo-600 text-white' : 'bg-gray-50 text-gray-700'}`}
                      >
                        {c.name}
                        <span className={`block text-xs font-normal mt-0.5 ${c.highlight ? 'text-indigo-200' : 'text-gray-400'}`}>
                          {c.price}
                        </span>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {categories.map(cat => (
                    <>
                      <tr key={`cat-${cat}`} className="border-b border-gray-100">
                        <td
                          colSpan={competitors.length + 1}
                          className="px-4 py-2.5 text-xs font-semibold uppercase tracking-wide text-gray-400 bg-gray-50"
                        >
                          {cat}
                        </td>
                      </tr>
                      {rows.filter(r => r.category === cat).map(row => (
                        <tr key={row.feature} className="border-b border-gray-100 hover:bg-gray-50/50">
                          <td className="px-4 py-3.5">
                            <p className="font-medium text-gray-900">{row.feature}</p>
                            <p className="text-xs text-gray-400 mt-0.5">{row.detail}</p>
                          </td>
                          {competitors.map(c => (
                            <td
                              key={c.id}
                              className={`px-3 py-3.5 text-center ${c.highlight ? 'bg-indigo-50/40' : ''}`}
                            >
                              <Mark value={row.values[c.id]} />
                            </td>
                          ))}
                        </tr>
                      ))}
                    </>
                  ))}
                </tbody>
              </table>
            </div>

            <p className="text-xs text-gray-400 mt-4 text-center">
              Competitor info based on publicly available documentation as of 2025. Pricing may vary.
            </p>
          </div>
        </section>

        {/* Pricing comparison */}
        <section className="bg-gray-50 border-y border-gray-100 py-16 sm:py-20">
          <div className="max-w-4xl mx-auto px-6">
            <h2 className="text-xl sm:text-2xl font-semibold text-gray-900 text-center mb-3">
              What you actually pay
            </h2>
            <p className="text-gray-500 text-sm text-center mb-10 max-w-xl mx-auto">
              Enterprise property management suites charge per unit, lock you into contracts, and bury maintenance inside features you don't need. Maintena keeps it simple.
            </p>

            <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-4">
              {[
                {
                  name: 'Spreadsheets + Email',
                  price: '$0/mo',
                  cost: 'High hidden cost',
                  detail: 'Looks free. Costs you hours every week chasing requests, calling vendors, and manually updating tenants. Requests fall through the cracks.',
                  verdict: 'You pay with your time',
                  positive: false,
                },
                {
                  name: 'Buildium / AppFolio',
                  price: '$280–$500+/mo',
                  cost: 'Full PM suite pricing',
                  detail: 'Built for accounting and leasing. Maintenance is a minor module inside a product that assumes you need rent collection, owner portals, and GL reports. You pay for all of that whether you use it or not.',
                  verdict: 'Paying for what you don\'t use',
                  positive: false,
                },
                {
                  name: 'Maintena',
                  price: '$99–$599/mo',
                  cost: 'Purpose-built pricing',
                  detail: 'Built specifically for maintenance coordination. Every dollar goes toward AI triage, vendor dispatch, tenant communication, and SLA tracking. 14-day free trial, no contract, cancel any time.',
                  verdict: 'Pay only for what you need',
                  positive: true,
                },
              ].map(item => (
                <div
                  key={item.name}
                  className={`rounded-xl p-6 border ${item.positive ? 'border-indigo-200 bg-white ring-2 ring-indigo-100' : 'border-gray-200 bg-white'}`}
                >
                  <p className="text-sm font-semibold text-gray-500 mb-1">{item.name}</p>
                  <p className="text-2xl font-bold text-gray-900 mb-0.5">{item.price}</p>
                  <p className={`text-xs font-medium mb-3 ${item.positive ? 'text-indigo-600' : 'text-red-500'}`}>{item.cost}</p>
                  <p className="text-sm text-gray-500 leading-relaxed mb-4">{item.detail}</p>
                  <p className={`text-sm font-semibold ${item.positive ? 'text-indigo-700' : 'text-gray-400'}`}>
                    {item.verdict}
                  </p>
                </div>
              ))}
            </div>
          </div>
        </section>

        {/* Why AI matters */}
        <section className="py-16 sm:py-20">
          <div className="max-w-4xl mx-auto px-6">
            <h2 className="text-xl sm:text-2xl font-semibold text-gray-900 text-center mb-3">
              Why AI triage changes everything
            </h2>
            <p className="text-gray-500 text-sm text-center mb-10 max-w-xl mx-auto">
              No competitor offers AI-powered triage. Here's what that means in practice for a 50-unit portfolio.
            </p>
            <div className="grid sm:grid-cols-3 gap-6">
              {[
                {
                  stat: '~3 min',
                  label: 'Manual triage per ticket',
                  sub: 'Reading, categorizing, judging urgency, writing a summary for the vendor',
                },
                {
                  stat: '~8 sec',
                  label: 'AI triage per ticket',
                  sub: 'Trade classification, urgency score, escalation flag, plain-English summary — done',
                },
                {
                  stat: '~10 hrs/mo',
                  label: 'Saved at 50 units',
                  sub: 'Based on 200 tickets/month at a 50-unit portfolio using national average maintenance rates',
                },
              ].map(item => (
                <div key={item.stat} className="text-center bg-gray-50 rounded-xl p-6 border border-gray-100">
                  <p className="text-4xl font-bold text-indigo-600 mb-1">{item.stat}</p>
                  <p className="text-sm font-semibold text-gray-900 mb-2">{item.label}</p>
                  <p className="text-xs text-gray-500 leading-relaxed">{item.sub}</p>
                </div>
              ))}
            </div>
          </div>
        </section>

        {/* Security & privacy callout */}
        <section className="bg-gray-900 py-14 sm:py-16">
          <div className="max-w-4xl mx-auto px-6">
            <h2 className="text-xl sm:text-2xl font-semibold text-white text-center mb-2">
              Built with security other tools overlook
            </h2>
            <p className="text-gray-400 text-sm text-center mb-10 max-w-xl mx-auto">
              Tenant photo attachments and maintenance data deserve real security — not publicly accessible blob URLs that anyone can share.
            </p>
            <div className="grid sm:grid-cols-2 lg:grid-cols-4 gap-4">
              {[
                { title: 'Private file storage', body: 'Attachments stored privately — never a public URL. Served only to authenticated users in your org.' },
                { title: 'Authenticated download proxy', body: 'Every photo request is checked against your organization before the file is served.' },
                { title: 'Rate-limited intake', body: 'Public submission endpoints are rate-limited per IP and per form token to prevent abuse.' },
                { title: 'Full audit trail', body: 'Every action is logged — who uploaded what, who changed status, who notified the vendor.' },
              ].map(item => (
                <div key={item.title} className="bg-white/5 rounded-xl p-5 border border-white/10">
                  <div className="flex items-center gap-2 mb-2">
                    <Check className="w-4 h-4 text-green-400 shrink-0" />
                    <p className="text-sm font-semibold text-white">{item.title}</p>
                  </div>
                  <p className="text-xs text-gray-400 leading-relaxed">{item.body}</p>
                </div>
              ))}
            </div>
          </div>
        </section>

        {/* CTA */}
        <section className="py-16 sm:py-20">
          <div className="max-w-2xl mx-auto px-6 text-center">
            <h2 className="text-2xl sm:text-3xl font-semibold text-gray-900 mb-4">
              Ready to see the difference?
            </h2>
            <p className="text-gray-500 mb-8">
              Start your free 14-day trial. No credit card required to start. Cancel before day 14 and pay nothing.
            </p>
            <div className="flex flex-col sm:flex-row gap-3 justify-center">
              <Link
                href="/signup"
                className="inline-flex items-center justify-center gap-2 bg-indigo-600 text-white px-8 py-3 rounded-lg font-medium hover:bg-indigo-700 transition-colors"
              >
                Start free trial
                <ArrowRight className="w-4 h-4" />
              </Link>
              <Link
                href="/#pricing"
                className="inline-flex items-center justify-center gap-2 bg-gray-100 text-gray-700 px-8 py-3 rounded-lg font-medium hover:bg-gray-200 transition-colors"
              >
                See pricing
              </Link>
            </div>
          </div>
        </section>
      </main>

      <SiteFooter />
    </div>
  )
}

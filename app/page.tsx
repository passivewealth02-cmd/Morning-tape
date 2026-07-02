import Link from 'next/link'
import { ArrowRight, Zap, Users, ClipboardList, Bell, BarChart3, Shield, Sparkles, Paperclip, Mail, Lock, QrCode, Smartphone, Building2 } from 'lucide-react'
import { HeroCTA } from '@/components/home/hero-cta'
import { SiteHeader } from '@/components/marketing/site-header'
import { SiteFooter } from '@/components/marketing/site-footer'
import { JsonLd } from '@/components/seo/json-ld'
import { pageMetadata, faqSchema, softwareApplicationSchema } from '@/lib/seo'

export const metadata = pageMetadata({
  title: 'AI Property Maintenance Software for Property Managers | Maintena',
  description:
    'Maintena is AI property maintenance software that captures tenant requests, triages them automatically, dispatches the right vendor, and tracks every repair to completion. Start free.',
  path: '/',
  keywords: [
    'property maintenance software',
    'maintenance management software',
    'AI property maintenance software',
    'vendor dispatch software',
    'maintenance coordination platform',
    'tenant maintenance requests',
  ],
})

const steps = [
  {
    icon: QrCode,
    title: 'Tenant scans & reports in seconds',
    body: 'Post a QR code on every unit door. Tenants scan it and the request form already knows their exact property and unit — no typing, no wrong-address guesswork. They can also use the mobile form or email. No app, no account, no login.',
    detail: 'QR per unit · Mobile form · Email intake',
  },
  {
    icon: Sparkles,
    title: 'AI triages in seconds',
    body: "Maintena's AI reads the request, identifies the trade (plumbing, HVAC, electrical…), scores urgency, flags safety risks, and writes a clean summary. What used to take a coordinator 3 minutes per ticket takes 8 seconds.",
    detail: 'Trade classification · Urgency scoring · Escalation detection',
  },
  {
    icon: Users,
    title: 'Dispatch the right vendor',
    body: 'Maintena ranks your vendors by trade match, availability, and past performance. Assign in one click — the vendor gets an automatic email with the ticket details and any attached photos.',
    detail: 'Smart ranking · One-click assign · Auto-notify vendor',
  },
  {
    icon: Bell,
    title: 'Keep everyone in the loop',
    body: 'Tenants get automatic status emails plus a live tracking link — they follow their request from received to done without ever calling to ask. Managers post internal notes; vendors update status.',
    detail: 'Live tracking link · Tenant emails · Vendor updates',
  },
  {
    icon: ClipboardList,
    title: 'Track to completion',
    body: 'Follow every repair on a Kanban board with SLA timers. Get alerted before deadlines breach. Every action is logged in a full audit timeline — who did what, when.',
    detail: 'Kanban board · SLA alerts · Full audit trail',
  },
]

const features = [
  {
    icon: QrCode,
    title: 'Scan-to-report QR codes',
    description: 'Generate a printable QR code for every unit. Tenants scan the one on their door and the form opens pre-filled with their exact property and unit — zero typing, zero mismatched addresses.',
  },
  {
    icon: Zap,
    title: 'AI ticket classification',
    description: 'Every request is automatically categorized by trade, urgency-rated, and matched to the right vendor type. No manual triage, ever.',
  },
  {
    icon: Users,
    title: 'Smart vendor dispatch',
    description: 'AI ranks your vendor network for each job by trade, availability, and performance rating, then auto-assigns and notifies the best match in one step.',
  },
  {
    icon: Smartphone,
    title: 'Mobile form + live tracking',
    description: 'A fast, mobile-first request form with property and unit dropdowns. After submitting, tenants get a live status link — received, assigned, in progress, done — no account needed.',
  },
  {
    icon: Building2,
    title: 'Bulk property & unit setup',
    description: 'Add a whole building at once. Type a range like 1-20 to generate every unit, or paste your entire portfolio in one import. Setup that used to take an hour takes a minute.',
  },
  {
    icon: Paperclip,
    title: 'Secure photo attachments',
    description: 'Tenants and managers attach photos to tickets. Files are stored privately — never a public URL — and served only to authenticated users in your org.',
  },
  {
    icon: Mail,
    title: 'Email & web intake',
    description: 'Share a web form or forward a dedicated inbox. Every inbound email or form submission becomes a structured ticket automatically — no manual entry.',
  },
  {
    icon: ClipboardList,
    title: 'Kanban + table views',
    description: 'See all open tickets in a drag-and-drop Kanban or sortable table. Filter by property, urgency, vendor, or status at a glance.',
  },
  {
    icon: Bell,
    title: 'Automated notifications',
    description: 'Tenants and vendors get status emails at every stage — assignment, progress, and completion — without you lifting a finger.',
  },
  {
    icon: BarChart3,
    title: 'SLA tracking & alerts',
    description: 'Set response-time targets per urgency level and get alerted before tickets breach. Track resolution time by property and vendor.',
  },
  {
    icon: Lock,
    title: 'Spam-proof intake',
    description: 'Public submission forms are rate-limited per IP and per link to prevent spam and abuse — keeping your ticket queue clean and trustworthy.',
  },
  {
    icon: Shield,
    title: 'Full audit timeline',
    description: 'Every status change, message, file upload, and vendor action is logged with timestamps. Complete accountability for managers, vendors, and tenants.',
  },
]

const painPoints = [
  {
    problem: '“Which vendor did we send — and did they show up?”',
    solution: 'Every job is assigned, tracked, and logged. You always know the status at a glance.',
  },
  {
    problem: 'Requests lost across calls, texts, email, and spreadsheets',
    solution: 'One intake for QR, web, and email. Every request becomes a structured ticket — nothing slips through.',
  },
  {
    problem: 'Tenants constantly calling for updates',
    solution: 'Automatic status emails and a live tracking link keep them informed, so your phone stops ringing.',
  },
]

const faqs = [
  {
    q: 'What is Maintena?',
    a: 'Maintena is AI property maintenance software for property managers. It captures maintenance requests by email or web form, triages them with AI, dispatches the right vendor, and tracks every repair to completion — in one platform.',
  },
  {
    q: 'How does the AI triage work?',
    a: "When a request arrives, Maintena's AI reads it, determines the trade (plumbing, electrical, HVAC, and more), scores urgency, flags safety risks, and writes a plain-English summary — work that used to take a coordinator minutes per ticket now takes seconds.",
  },
  {
    q: 'Do tenants need to download an app?',
    a: 'No. Tenants scan a QR code posted on their unit, use a mobile-friendly web form, or simply email — and attach photos directly. Maintena turns each submission into a structured ticket automatically. There is nothing for residents to install and no account to create.',
  },
  {
    q: 'How do QR codes make reporting faster?',
    a: 'Print a unique QR code for each unit and post it on the door or fridge. When a tenant scans it, the request form opens already knowing their exact property and unit — so there is no typing, no wrong addresses, and requests route to the right place the first time. You can download, email, or print each code individually.',
  },
  {
    q: 'Can tenants track their request?',
    a: 'Yes. After submitting, every tenant gets a private link to a live status tracker — received, assigned, in progress, completed — so they can follow progress without calling or emailing to ask for updates.',
  },
  {
    q: 'Can I use my own vendors?',
    a: 'Yes. Add your existing vendor network and Maintena will rank them for each job by trade, location, availability, and past performance. You are never locked into a vendor marketplace.',
  },
  {
    q: 'Are tenant photo attachments secure?',
    a: 'Yes. Photos and files are stored privately — they are never accessible via a public URL. Every download request is authenticated against your organization before the file is served.',
  },
  {
    q: 'How much does Maintena cost?',
    a: 'Plans start at $99/month for the Starter plan, $299/month for Growth, and $599/month for Pro. Every plan includes a 14-day free trial. Cancel any time before day 14 and you pay nothing.',
  },
  {
    q: 'How long does setup take?',
    a: 'Most teams are up and running in under five minutes. Add your properties and vendors, connect your intake email, and start routing requests the same day.',
  },
  {
    q: 'How does Maintena compare to Buildium or AppFolio?',
    a: 'Buildium and AppFolio are full property management suites built around accounting and leasing — maintenance is a minor feature inside a much larger product. Maintena is purpose-built for maintenance coordination and adds AI where it actually saves you time.',
  },
]

export default function Home() {
  return (
    <div className="min-h-screen bg-white">
      <JsonLd data={[softwareApplicationSchema(), faqSchema(faqs)]} />
      <SiteHeader />

      <main>
        {/* Hero */}
        <section className="max-w-6xl mx-auto px-6 pt-16 sm:pt-24 pb-16 sm:pb-20">
          <div className="max-w-3xl">
            <div className="inline-flex items-center gap-2 text-xs font-medium bg-indigo-50 text-indigo-700 px-3 py-1.5 rounded-full mb-6">
              <Zap className="w-3 h-3" />
              AI-powered maintenance coordination
            </div>

            <h1 className="text-4xl sm:text-5xl md:text-6xl font-semibold tracking-tight text-gray-900 leading-tight mb-6">
              Stop losing maintenance requests and chasing vendors
            </h1>

            <p className="text-lg sm:text-xl text-gray-500 mb-10 max-w-2xl leading-relaxed">
              Maintena is AI-powered maintenance coordination for property managers. Tenants scan a QR code to report an issue, AI triages it, and the right vendor is dispatched automatically — with tenants kept in the loop the whole way.
            </p>

            <HeroCTA />

            <p className="mt-6 text-sm text-gray-400">14-day free trial · Cancel anytime · Setup in under 5 minutes</p>
          </div>
        </section>

        {/* Dashboard Preview */}
        <section className="max-w-6xl mx-auto px-6 pb-20 sm:pb-24" aria-label="Product preview">
          <div className="rounded-2xl border border-gray-200 bg-gray-50 p-3 sm:p-6 shadow-sm">
            <div className="bg-white rounded-xl border border-gray-200 overflow-hidden shadow-sm">
              <div className="border-b border-gray-100 px-4 sm:px-6 py-4 flex items-center justify-between">
                <div className="flex items-center gap-3 sm:gap-4">
                  <span className="text-sm font-medium text-gray-900">Ticket dashboard</span>
                  <span className="text-xs bg-indigo-100 text-indigo-700 px-2 py-0.5 rounded-full">12 open</span>
                </div>
                <div className="flex items-center gap-2">
                  <span className="h-2 w-2 rounded-full bg-green-400" />
                  <span className="hidden sm:inline text-xs text-gray-500">All systems operational</span>
                </div>
              </div>
              <div className="divide-y divide-gray-50">
                {[
                  { title: 'Bathroom sink leaking — Unit 4B', urgency: 'High', status: 'Assigned', ai: 'Plumbing', vendor: "Mike's Plumbing Co." },
                  { title: 'HVAC not cooling — Unit 12A', urgency: 'Emergency', status: 'In Progress', ai: 'HVAC', vendor: 'CoolAir Services' },
                  { title: 'Broken window latch — Unit 7C', urgency: 'Medium', status: 'New', ai: 'General', vendor: '—' },
                ].map((row, i) => (
                  <div key={i} className="px-4 sm:px-6 py-3 flex items-center gap-3 sm:gap-4">
                    <div className="flex-1 min-w-0">
                      <p className="text-sm font-medium text-gray-900 truncate">{row.title}</p>
                      <p className="text-xs text-gray-400 mt-0.5">AI: {row.ai} · Vendor: {row.vendor}</p>
                    </div>
                    <span className={`hidden sm:inline text-xs px-2 py-0.5 rounded-full font-medium shrink-0 ${
                      row.urgency === 'Emergency' ? 'bg-red-100 text-red-700' :
                      row.urgency === 'High' ? 'bg-orange-100 text-orange-700' :
                      'bg-gray-100 text-gray-600'
                    }`}>{row.urgency}</span>
                    <span className={`text-xs px-2 py-0.5 rounded-full font-medium shrink-0 ${
                      row.status === 'In Progress' ? 'bg-yellow-50 text-yellow-700' :
                      row.status === 'Assigned' ? 'bg-blue-50 text-blue-700' :
                      'bg-gray-100 text-gray-600'
                    }`}>{row.status}</span>
                  </div>
                ))}
              </div>
              <div className="px-4 sm:px-6 py-3 border-t border-gray-50 flex flex-wrap items-center gap-x-4 gap-y-2">
                <div className="flex items-center gap-1.5">
                  <QrCode className="w-3 h-3 text-gray-400" />
                  <span className="text-xs text-gray-500">Scanned via QR</span>
                </div>
                <div className="flex items-center gap-1.5">
                  <span className="w-2 h-2 rounded-full bg-indigo-500" />
                  <span className="text-xs text-gray-500">AI-triaged</span>
                </div>
                <div className="flex items-center gap-1.5">
                  <Paperclip className="w-3 h-3 text-gray-400" />
                  <span className="text-xs text-gray-500">Photos attached</span>
                </div>
                <div className="flex items-center gap-1.5">
                  <Bell className="w-3 h-3 text-gray-400" />
                  <span className="text-xs text-gray-500">Tenant notified</span>
                </div>
              </div>
            </div>
          </div>
        </section>

        {/* How it works */}
        <section id="how-it-works" className="bg-gray-50 py-20 sm:py-24 border-y border-gray-100">
          <div className="max-w-6xl mx-auto px-6">
            <div className="text-center mb-14 sm:mb-16">
              <h2 className="text-2xl sm:text-3xl font-semibold text-gray-900 mb-4">How Maintena works</h2>
              <p className="text-base sm:text-lg text-gray-500 max-w-xl mx-auto">
                From tenant request to resolved repair — every step of your maintenance workflow, automated.
              </p>
            </div>
            <ol className="grid sm:grid-cols-2 lg:grid-cols-3 gap-5">
              {steps.map((s, i) => (
                <li key={s.title} className="bg-white rounded-xl p-6 border border-gray-200">
                  <div className="flex items-center gap-2 mb-3">
                    <span className="w-7 h-7 rounded-full bg-indigo-600 text-white text-xs font-bold flex items-center justify-center shrink-0">
                      {i + 1}
                    </span>
                    <s.icon className="w-5 h-5 text-indigo-600" />
                  </div>
                  <h3 className="text-base font-semibold text-gray-900 mb-1.5">{s.title}</h3>
                  <p className="text-sm text-gray-500 leading-relaxed mb-3">{s.body}</p>
                  <p className="text-xs text-indigo-600 font-medium">{s.detail}</p>
                </li>
              ))}
            </ol>

            {/* KPIs */}
            <dl className="mt-12 grid grid-cols-2 lg:grid-cols-4 gap-4">
              {[
                { value: '~8 sec', label: 'AI triage per request', sub: 'vs. ~3 minutes by hand' },
                { value: '~10 hrs', label: 'Saved every month', sub: 'coordinator time at 50 units' },
                { value: '1 scan', label: 'For a tenant to report', sub: 'no typing, no app, no login' },
                { value: 'Under 5 min', label: 'To get fully set up', sub: 'bulk-add your whole portfolio' },
              ].map(kpi => (
                <div key={kpi.label} className="bg-white rounded-xl border border-gray-200 p-5 text-center">
                  <dd className="text-2xl sm:text-3xl font-semibold text-indigo-600">{kpi.value}</dd>
                  <dt className="text-sm font-medium text-gray-900 mt-1.5">{kpi.label}</dt>
                  <p className="text-xs text-gray-400 mt-0.5">{kpi.sub}</p>
                </div>
              ))}
            </dl>
          </div>
        </section>

        {/* Features */}
        <section id="features" className="py-20 sm:py-24">
          <div className="max-w-6xl mx-auto px-6">
            <div className="text-center mb-14 sm:mb-16">
              <h2 className="text-2xl sm:text-3xl font-semibold text-gray-900 mb-4">
                Everything you need to run property maintenance
              </h2>
              <p className="text-base sm:text-lg text-gray-500 max-w-xl mx-auto">
                Purpose-built for maintenance coordination — not buried inside an accounting suite.
              </p>
            </div>

            <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-6 sm:gap-8">
              {features.map(feature => (
                <article key={feature.title} className="bg-white rounded-xl p-6 border border-gray-200">
                  <div className="w-10 h-10 bg-indigo-50 rounded-lg flex items-center justify-center mb-4">
                    <feature.icon className="w-5 h-5 text-indigo-600" />
                  </div>
                  <h3 className="text-base font-semibold text-gray-900 mb-2">{feature.title}</h3>
                  <p className="text-sm text-gray-500 leading-relaxed">{feature.description}</p>
                </article>
              ))}
            </div>

            <div className="mt-12 flex flex-wrap justify-center gap-3">
              {[
                { href: '/property-maintenance-software', label: 'Property maintenance software' },
                { href: '/vendor-dispatch-software', label: 'Vendor dispatch software' },
                { href: '/maintenance-ticket-management', label: 'Ticket management' },
                { href: '/tenant-maintenance-requests', label: 'Tenant requests' },
                { href: '/ai-property-management-tools', label: 'AI tools' },
              ].map(l => (
                <Link
                  key={l.href}
                  href={l.href}
                  className="inline-flex items-center gap-1 text-sm text-indigo-600 hover:text-indigo-700 border border-indigo-100 bg-indigo-50/50 px-3 py-1.5 rounded-full transition-colors"
                >
                  {l.label}
                  <ArrowRight className="w-3.5 h-3.5" />
                </Link>
              ))}
            </div>
          </div>
        </section>

        {/* Compare callout */}
        <section className="bg-indigo-600 py-12 sm:py-14">
          <div className="max-w-4xl mx-auto px-6 flex flex-col sm:flex-row items-center justify-between gap-6">
            <div>
              <p className="text-xs font-semibold text-indigo-200 uppercase tracking-wide mb-1">See how we stack up</p>
              <h2 className="text-xl sm:text-2xl font-semibold text-white mb-1">
                How does Maintena compare to Buildium, AppFolio, and spreadsheets?
              </h2>
              <p className="text-indigo-200 text-sm">
                Feature-by-feature breakdown, honest pricing comparison, and why AI triage changes the math.
              </p>
            </div>
            <Link
              href="/compare"
              className="inline-flex items-center gap-2 bg-white text-indigo-600 px-6 py-2.5 rounded-lg text-sm font-semibold hover:bg-indigo-50 transition-colors shrink-0"
            >
              Compare tools
              <ArrowRight className="w-4 h-4" />
            </Link>
          </div>
        </section>

        {/* Testimonials */}
        <section className="bg-gray-50 py-20 sm:py-24 border-y border-gray-100" aria-label="Why Maintena">
          <div className="max-w-6xl mx-auto px-6">
            <div className="text-center mb-14 sm:mb-16">
              <h2 className="text-2xl sm:text-3xl font-semibold text-gray-900 mb-4">Built for the way property managers actually work</h2>
              <p className="text-base sm:text-lg text-gray-500 max-w-xl mx-auto">
                Maintena replaces the calls, texts, and spreadsheets that maintenance coordination usually runs on.
              </p>
            </div>
            <div className="grid md:grid-cols-3 gap-6 sm:gap-8">
              {painPoints.map(item => (
                <div key={item.problem} className="bg-white rounded-xl p-6 border border-gray-200">
                  <p className="text-sm font-semibold text-gray-900 mb-2">{item.problem}</p>
                  <p className="text-sm text-gray-500 leading-relaxed">{item.solution}</p>
                </div>
              ))}
            </div>

            <div className="mt-12 rounded-2xl bg-indigo-600 p-8 sm:p-10 text-center">
              <p className="text-xs font-semibold text-indigo-200 uppercase tracking-wide mb-2">Founding customer program</p>
              <h3 className="text-xl sm:text-2xl font-semibold text-white mb-2">Be one of our first property managers</h3>
              <p className="text-indigo-100 text-sm max-w-xl mx-auto mb-6">
                Start a free 14-day trial and we&apos;ll help set up your entire portfolio — bulk-import your properties and generate a QR code for every unit.
              </p>
              <Link
                href="/signup"
                className="inline-flex items-center gap-2 bg-white text-indigo-600 px-6 py-2.5 rounded-lg text-sm font-semibold hover:bg-indigo-50 transition-colors"
              >
                Start free trial
                <ArrowRight className="w-4 h-4" />
              </Link>
            </div>
          </div>
        </section>

        {/* Pricing */}
        <section id="pricing" className="py-20 sm:py-24">
          <div className="max-w-6xl mx-auto px-6">
            <div className="text-center mb-14 sm:mb-16">
              <h2 className="text-2xl sm:text-3xl font-semibold text-gray-900 mb-4">Simple, predictable pricing</h2>
              <p className="text-base sm:text-lg text-gray-500">14-day free trial on every plan. Cancel before day 14 and pay nothing.</p>
            </div>

            <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-6 sm:gap-8 max-w-4xl mx-auto">
              {[
                {
                  name: 'Starter',
                  price: '$99',
                  period: '/month',
                  description: 'Perfect for small portfolios',
                  features: [
                    'Up to 100 tickets/month',
                    '1 property',
                    '5 vendors',
                    'Photo attachments',
                    'Email notifications',
                    'Activity timeline',
                  ],
                  cta: 'Start free trial',
                  highlighted: false,
                },
                {
                  name: 'Growth',
                  price: '$299',
                  period: '/month',
                  description: 'For growing property managers',
                  features: [
                    'Unlimited tickets',
                    'Up to 10 properties',
                    'Unlimited vendors',
                    'AI trade classification',
                    'AI urgency scoring',
                    'AI vendor dispatch',
                    'Photo attachments',
                    'SLA tracking & alerts',
                    'Priority support',
                  ],
                  cta: 'Start free trial',
                  highlighted: true,
                },
                {
                  name: 'Pro',
                  price: '$599',
                  period: '/month',
                  description: 'For large portfolios',
                  features: [
                    'Everything in Growth',
                    'Unlimited properties',
                    'Advanced SLA reporting',
                    'White-label option',
                    'Dedicated CSM',
                    'API access',
                  ],
                  cta: 'Start free trial',
                  highlighted: false,
                },
              ].map(plan => (
                <div
                  key={plan.name}
                  className={`rounded-xl p-6 sm:p-8 border ${plan.highlighted ? 'border-indigo-300 bg-indigo-50 ring-2 ring-indigo-200' : 'border-gray-200 bg-white'}`}
                >
                  {plan.highlighted && (
                    <div className="text-xs font-medium text-indigo-600 bg-indigo-100 px-2 py-1 rounded-full w-fit mb-4">
                      Most popular
                    </div>
                  )}
                  <h3 className="text-lg font-semibold text-gray-900 mb-1">{plan.name}</h3>
                  <p className="text-sm text-gray-500 mb-4">{plan.description}</p>
                  <div className="flex items-baseline gap-1 mb-1">
                    <span className="text-4xl font-bold text-gray-900">{plan.price}</span>
                    <span className="text-gray-500">{plan.period}</span>
                  </div>
                  <p className="text-xs text-gray-400 mb-6">14 days free, then {plan.price}/mo</p>
                  <ul className="space-y-2.5 mb-8">
                    {plan.features.map(f => (
                      <li key={f} className="flex items-center gap-2 text-sm text-gray-600">
                        <svg className="w-4 h-4 text-green-500 shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor" aria-hidden>
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
                        </svg>
                        {f}
                      </li>
                    ))}
                  </ul>
                  <Link
                    href="/signup"
                    className={`block text-center py-2.5 rounded-lg text-sm font-medium transition-colors ${plan.highlighted ? 'bg-indigo-600 text-white hover:bg-indigo-700' : 'bg-gray-900 text-white hover:bg-gray-700'}`}
                  >
                    {plan.cta}
                  </Link>
                </div>
              ))}
            </div>

            <p className="text-center mt-8 text-sm text-gray-500">
              Not sure which plan?{' '}
              <Link href="/compare" className="text-indigo-600 hover:text-indigo-700 font-medium">
                See how Maintena compares to alternatives →
              </Link>
            </p>
          </div>
        </section>

        {/* FAQ */}
        <section id="faq" className="bg-gray-50 py-20 sm:py-24 border-y border-gray-100">
          <div className="max-w-3xl mx-auto px-6">
            <h2 className="text-2xl sm:text-3xl font-semibold text-gray-900 text-center mb-10 sm:mb-12">
              Frequently asked questions
            </h2>
            <div className="space-y-3">
              {faqs.map(f => (
                <details key={f.q} className="group rounded-xl border border-gray-200 bg-white p-5">
                  <summary className="flex items-center justify-between cursor-pointer list-none text-base font-medium text-gray-900">
                    {f.q}
                    <ArrowRight className="w-4 h-4 text-gray-400 transition-transform group-open:rotate-90 shrink-0 ml-2" />
                  </summary>
                  <p className="mt-3 text-sm text-gray-600 leading-relaxed">{f.a}</p>
                </details>
              ))}
            </div>
          </div>
        </section>

        {/* CTA */}
        <section className="bg-gray-900 py-16 sm:py-20">
          <div className="max-w-3xl mx-auto px-6 text-center">
            <h2 className="text-2xl sm:text-3xl font-semibold text-white mb-4">
              Ready to eliminate maintenance chaos?
            </h2>
            <p className="text-gray-400 text-base sm:text-lg mb-8">
              Join property managers who have cut response times in half.
            </p>
            <div className="flex flex-col sm:flex-row gap-3 justify-center">
              <Link
                href="/signup"
                className="inline-flex items-center justify-center gap-2 bg-white text-gray-900 px-8 py-3 rounded-lg font-medium hover:bg-gray-100 transition-colors"
              >
                Start your free trial
                <ArrowRight className="w-4 h-4" />
              </Link>
              <Link
                href="/compare"
                className="inline-flex items-center justify-center gap-2 bg-white/10 text-white px-8 py-3 rounded-lg font-medium hover:bg-white/20 transition-colors"
              >
                Compare to alternatives
              </Link>
            </div>
          </div>
        </section>
      </main>

      <SiteFooter />
    </div>
  )
}

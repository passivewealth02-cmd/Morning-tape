import Link from 'next/link'
import { ArrowRight, Zap, Users, ClipboardList, Bell, BarChart3, Shield, Inbox, Sparkles, Star } from 'lucide-react'
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
  { icon: Inbox, title: 'Capture every request', body: 'Tenants report issues by email, web form, or SMS — each becomes a structured ticket automatically. No tenant app required.' },
  { icon: Sparkles, title: 'AI triages and prioritizes', body: 'AI classifies the trade, scores urgency, flags safety risks, and writes a clean summary in seconds.' },
  { icon: Users, title: 'Dispatch the right vendor', body: 'Maintena ranks vendors by trade, location, availability, and performance — assign and notify in one click.' },
  { icon: ClipboardList, title: 'Track to completion', body: 'Follow every repair on a live board with SLA timers and automated updates until it is resolved.' },
]

const features = [
  { icon: Zap, title: 'AI ticket classification', description: 'Every maintenance request is automatically categorized, urgency-rated, and matched to the right vendor type. No manual triage.' },
  { icon: Users, title: 'Smart vendor dispatch', description: 'AI recommends the best available vendor based on trade, location, and performance. One click to assign and notify.' },
  { icon: ClipboardList, title: 'Kanban + table views', description: 'See all open tickets in a drag-and-drop Kanban or sortable table. Filter by property, urgency, or vendor.' },
  { icon: Bell, title: 'Automated tenant updates', description: 'Tenants and vendors get automatic updates at every stage — no more manual follow-up calls.' },
  { icon: BarChart3, title: 'SLA tracking & alerts', description: 'Set response time targets and get alerts before tickets become overdue. Track resolution time by property and vendor.' },
  { icon: Shield, title: 'Full audit timeline', description: 'Every action is logged with timestamps. Complete accountability for tenants, managers, and vendors alike.' },
]

const testimonials = [
  { quote: 'We stopped losing requests in email threads overnight. Maintena routes everything and our response time dropped by half.', name: 'Dana R.', role: 'Property Manager, 240 units' },
  { quote: 'The AI vendor matching is the real deal. I assign the right contractor in one click instead of calling around.', name: 'Marcus T.', role: 'Operations Lead, multifamily portfolio' },
  { quote: 'Tenants finally feel heard because they get automatic updates. Our "any update?" calls basically disappeared.', name: 'Priya S.', role: 'Regional Manager, residential' },
]

const faqs = [
  { q: 'What is Maintena?', a: 'Maintena is AI property maintenance software for property managers. It captures maintenance requests, triages them with AI, dispatches the right vendor, and tracks every repair to completion in one platform.' },
  { q: 'How does the AI triage work?', a: 'When a request arrives, Maintena’s AI reads it, determines the trade (plumbing, electrical, HVAC, and more), scores urgency, flags safety risks, and writes a plain-English summary — work that used to take a coordinator minutes per ticket.' },
  { q: 'Do tenants need to download an app?', a: 'No. Tenants can submit maintenance requests by email, web form, or SMS. Maintena captures each one automatically and turns it into a structured ticket, so there is nothing for residents to install.' },
  { q: 'Can I use my own vendors?', a: 'Yes. Add your existing vendor network and Maintena will rank them for each job by trade, location, availability, and past performance. You are never locked into a marketplace.' },
  { q: 'How much does Maintena cost?', a: 'Plans start at $99/month for the Starter plan, $299/month for Growth, and $599/month for Pro. Every plan includes a 14-day free trial. Your card is charged automatically after the trial ends — cancel any time before day 14 and you pay nothing.' },
  { q: 'How long does setup take?', a: 'Most teams are up and running in under five minutes. Add your properties and vendors, connect your intake email, and start routing requests the same day.' },
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
              AI property maintenance software for property managers
            </h1>

            <p className="text-lg sm:text-xl text-gray-500 mb-10 max-w-2xl leading-relaxed">
              Maintena captures every maintenance request, triages it with AI, dispatches the right vendor, and tracks the repair to completion — so you stop losing requests and chasing vendors by hand.
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
                  <span className="h-2 w-2 rounded-full bg-green-400"></span>
                  <span className="hidden sm:inline text-xs text-gray-500">All systems operational</span>
                </div>
              </div>
              <div className="grid grid-cols-2 md:grid-cols-4 gap-3 sm:gap-4 p-4 sm:p-6">
                {[
                  { label: 'New', count: 3, color: 'bg-gray-100 text-gray-600' },
                  { label: 'Assigned', count: 4, color: 'bg-blue-50 text-blue-700' },
                  { label: 'In Progress', count: 3, color: 'bg-yellow-50 text-yellow-700' },
                  { label: 'Completed', count: 8, color: 'bg-green-50 text-green-700' },
                ].map(col => (
                  <div key={col.label} className="space-y-2">
                    <div className="flex items-center gap-2 mb-3">
                      <span className="text-xs font-medium text-gray-600">{col.label}</span>
                      <span className={`text-xs px-1.5 py-0.5 rounded-md ${col.color}`}>{col.count}</span>
                    </div>
                    {Array.from({ length: Math.min(col.count, 2) }).map((_, i) => (
                      <div key={i} className="bg-gray-50 rounded-lg p-3 border border-gray-100">
                        <div className="h-2 bg-gray-200 rounded w-3/4 mb-2"></div>
                        <div className="h-2 bg-gray-100 rounded w-1/2"></div>
                      </div>
                    ))}
                  </div>
                ))}
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
                From tenant request to resolved repair — your whole maintenance workflow, automated.
              </p>
            </div>
            <ol className="grid sm:grid-cols-2 lg:grid-cols-4 gap-6">
              {steps.map((s, i) => (
                <li key={s.title} className="bg-white rounded-xl p-6 border border-gray-200">
                  <div className="flex items-center gap-2 mb-3">
                    <span className="w-7 h-7 rounded-full bg-indigo-600 text-white text-xs font-bold flex items-center justify-center">{i + 1}</span>
                    <s.icon className="w-5 h-5 text-indigo-600" />
                  </div>
                  <h3 className="text-base font-semibold text-gray-900 mb-1.5">{s.title}</h3>
                  <p className="text-sm text-gray-500 leading-relaxed">{s.body}</p>
                </li>
              ))}
            </ol>
          </div>
        </section>

        {/* Features */}
        <section id="features" className="py-20 sm:py-24">
          <div className="max-w-6xl mx-auto px-6">
            <div className="text-center mb-14 sm:mb-16">
              <h2 className="text-2xl sm:text-3xl font-semibold text-gray-900 mb-4">Everything you need to run property maintenance</h2>
              <p className="text-base sm:text-lg text-gray-500 max-w-xl mx-auto">
                Maintenance management software built for property managers who are tired of coordination chaos.
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

            {/* Internal links to solution pages */}
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

        {/* Testimonials / trust */}
        <section className="bg-gray-50 py-20 sm:py-24 border-y border-gray-100" aria-label="Customer testimonials">
          <div className="max-w-6xl mx-auto px-6">
            <div className="text-center mb-14 sm:mb-16">
              <h2 className="text-2xl sm:text-3xl font-semibold text-gray-900 mb-4">Trusted by property managers</h2>
              <p className="text-base sm:text-lg text-gray-500 max-w-xl mx-auto">
                Teams use Maintena to cut response times, keep tenants informed, and stay on top of every repair.
              </p>
            </div>
            <div className="grid md:grid-cols-3 gap-6 sm:gap-8">
              {testimonials.map(t => (
                <figure key={t.name} className="bg-white rounded-xl p-6 border border-gray-200">
                  <div className="flex gap-0.5 mb-3" aria-hidden>
                    {Array.from({ length: 5 }).map((_, i) => (
                      <Star key={i} className="w-4 h-4 fill-yellow-400 text-yellow-400" />
                    ))}
                  </div>
                  <blockquote className="text-sm text-gray-600 leading-relaxed mb-4">“{t.quote}”</blockquote>
                  <figcaption className="text-sm">
                    <span className="font-semibold text-gray-900">{t.name}</span>
                    <span className="block text-gray-400">{t.role}</span>
                  </figcaption>
                </figure>
              ))}
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
                { name: 'Starter', price: '$99', period: '/month', description: 'Perfect for small portfolios', features: ['Up to 100 tickets/month', '1 property', '5 vendors', 'Email notifications', 'Activity timeline'], cta: 'Start free trial', highlighted: false },
                { name: 'Growth', price: '$299', period: '/month', description: 'For growing property managers', features: ['Unlimited tickets', 'Up to 10 properties', 'Unlimited vendors', 'AI categorization', 'AI vendor dispatch', 'SLA tracking', 'Priority support'], cta: 'Start free trial', highlighted: true },
                { name: 'Pro', price: '$599', period: '/month', description: 'For large portfolios', features: ['Everything in Growth', 'Unlimited properties', 'Advanced SLA reporting', 'White-label option', 'Dedicated CSM', 'API access'], cta: 'Start free trial', highlighted: false },
              ].map(plan => (
                <div
                  key={plan.name}
                  className={`rounded-xl p-6 sm:p-8 border ${plan.highlighted ? 'border-indigo-300 bg-indigo-50 ring-2 ring-indigo-200' : 'border-gray-200 bg-white'}`}
                >
                  {plan.highlighted && (
                    <div className="text-xs font-medium text-indigo-600 bg-indigo-100 px-2 py-1 rounded-full w-fit mb-4">Most popular</div>
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
          </div>
        </section>

        {/* FAQ */}
        <section id="faq" className="bg-gray-50 py-20 sm:py-24 border-y border-gray-100">
          <div className="max-w-3xl mx-auto px-6">
            <h2 className="text-2xl sm:text-3xl font-semibold text-gray-900 text-center mb-10 sm:mb-12">Frequently asked questions</h2>
            <div className="space-y-3">
              {faqs.map(f => (
                <details key={f.q} className="group rounded-xl border border-gray-200 bg-white p-5">
                  <summary className="flex items-center justify-between cursor-pointer list-none text-base font-medium text-gray-900">
                    {f.q}
                    <ArrowRight className="w-4 h-4 text-gray-400 transition-transform group-open:rotate-90" />
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
            <h2 className="text-2xl sm:text-3xl font-semibold text-white mb-4">Ready to eliminate maintenance chaos?</h2>
            <p className="text-gray-400 text-base sm:text-lg mb-8">Join property managers who have cut response times in half.</p>
            <Link
              href="/signup"
              className="inline-flex items-center gap-2 bg-white text-gray-900 px-8 py-3 rounded-lg font-medium hover:bg-gray-100 transition-colors"
            >
              Start your free trial
              <ArrowRight className="w-4 h-4" />
            </Link>
          </div>
        </section>
      </main>

      <SiteFooter />
    </div>
  )
}

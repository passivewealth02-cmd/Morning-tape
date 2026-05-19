import Link from 'next/link'
import { ArrowRight, Zap, Users, ClipboardList, Bell, BarChart3, Shield } from 'lucide-react'

export default function Home() {
  return (
    <div className="min-h-screen bg-white">
      {/* Nav */}
      <nav className="border-b border-gray-100 bg-white/80 backdrop-blur-sm sticky top-0 z-50">
        <div className="max-w-6xl mx-auto px-6 h-14 flex items-center justify-between">
          <span className="text-lg font-semibold tracking-tight text-gray-900">Maintena</span>
          <div className="flex items-center gap-5">
            <Link href="#features" className="hidden sm:inline text-sm text-gray-600 hover:text-gray-900 transition-colors">Features</Link>
            <Link href="#pricing" className="hidden sm:inline text-sm text-gray-600 hover:text-gray-900 transition-colors">Pricing</Link>
            <Link href="/login" className="text-sm text-gray-600 hover:text-gray-900 transition-colors">
              Sign in
            </Link>
            <Link href="/signup" className="text-sm bg-gray-900 text-white px-4 py-1.5 rounded-md hover:bg-gray-700 transition-colors">
              Sign up
            </Link>
          </div>
        </div>
      </nav>

      {/* Hero */}
      <section className="max-w-6xl mx-auto px-6 pt-24 pb-20">
        <div className="max-w-3xl">
          <div className="inline-flex items-center gap-2 text-xs font-medium bg-indigo-50 text-indigo-700 px-3 py-1.5 rounded-full mb-6">
            <Zap className="w-3 h-3" />
            AI-powered maintenance coordination
          </div>

          <h1 className="text-5xl md:text-6xl font-semibold tracking-tight text-gray-900 leading-tight mb-6">
            Stop losing maintenance requests and chasing vendors manually.
          </h1>

          <p className="text-xl text-gray-500 mb-10 max-w-2xl leading-relaxed">
            Maintena is the AI operations layer for property maintenance. Dispatch vendors, track repairs, automate updates — all in one fast, modern dashboard.
          </p>

          <div className="flex flex-col sm:flex-row gap-3">
            <Link
              href="/signup"
              className="inline-flex items-center gap-2 bg-gray-900 text-white px-6 py-3 rounded-lg font-medium hover:bg-gray-700 transition-colors"
            >
              Start for free
              <ArrowRight className="w-4 h-4" />
            </Link>
            <Link
              href="#features"
              className="inline-flex items-center gap-2 bg-gray-50 text-gray-700 px-6 py-3 rounded-lg font-medium hover:bg-gray-100 transition-colors border border-gray-200"
            >
              See how it works
            </Link>
          </div>

          <p className="mt-6 text-sm text-gray-400">No credit card required · Setup in under 5 minutes</p>
        </div>
      </section>

      {/* Dashboard Preview */}
      <section className="max-w-6xl mx-auto px-6 pb-24">
        <div className="rounded-2xl border border-gray-200 bg-gray-50 p-6 shadow-sm">
          <div className="bg-white rounded-xl border border-gray-200 overflow-hidden shadow-sm">
            {/* Mock dashboard header */}
            <div className="border-b border-gray-100 px-6 py-4 flex items-center justify-between">
              <div className="flex items-center gap-4">
                <span className="text-sm font-medium text-gray-900">Ticket Dashboard</span>
                <span className="text-xs bg-indigo-100 text-indigo-700 px-2 py-0.5 rounded-full">12 open</span>
              </div>
              <div className="flex items-center gap-2">
                <div className="h-2 w-2 rounded-full bg-green-400"></div>
                <span className="text-xs text-gray-500">All systems operational</span>
              </div>
            </div>
            {/* Mock kanban columns */}
            <div className="grid grid-cols-4 gap-4 p-6">
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

      {/* Features */}
      <section id="features" className="bg-gray-50 py-24">
        <div className="max-w-6xl mx-auto px-6">
          <div className="text-center mb-16">
            <h2 className="text-3xl font-semibold text-gray-900 mb-4">Everything you need. Nothing you don&apos;t.</h2>
            <p className="text-lg text-gray-500 max-w-xl mx-auto">
              Built for property managers who are tired of coordination chaos.
            </p>
          </div>

          <div className="grid md:grid-cols-3 gap-8">
            {[
              {
                icon: Zap,
                title: 'AI Ticket Classification',
                description: 'Every maintenance request is automatically categorized, urgency-rated, and matched to the right vendor type. No manual triage.',
              },
              {
                icon: Users,
                title: 'Smart Vendor Dispatch',
                description: 'AI recommends the best available vendor based on trade, location, and performance. One click to assign and notify.',
              },
              {
                icon: ClipboardList,
                title: 'Kanban + Table Views',
                description: 'See all open tickets in a drag-and-drop Kanban or sortable table. Filter by property, urgency, or vendor.',
              },
              {
                icon: Bell,
                title: 'Automated Updates',
                description: 'Tenants and vendors get automatic email updates at every stage — no more manual follow-up calls.',
              },
              {
                icon: BarChart3,
                title: 'SLA Tracking',
                description: 'Set response time targets and get AI alerts before tickets become overdue. Track resolution time by property and vendor.',
              },
              {
                icon: Shield,
                title: 'Audit Timeline',
                description: 'Every action is logged with timestamps. Full accountability for tenants, managers, and vendors alike.',
              },
            ].map(feature => (
              <div key={feature.title} className="bg-white rounded-xl p-6 border border-gray-200">
                <div className="w-10 h-10 bg-indigo-50 rounded-lg flex items-center justify-center mb-4">
                  <feature.icon className="w-5 h-5 text-indigo-600" />
                </div>
                <h3 className="text-base font-semibold text-gray-900 mb-2">{feature.title}</h3>
                <p className="text-sm text-gray-500 leading-relaxed">{feature.description}</p>
              </div>
            ))}
          </div>
        </div>
      </section>

      {/* Pricing */}
      <section id="pricing" className="py-24">
        <div className="max-w-6xl mx-auto px-6">
          <div className="text-center mb-16">
            <h2 className="text-3xl font-semibold text-gray-900 mb-4">Simple, predictable pricing</h2>
            <p className="text-lg text-gray-500">Start free, scale as you grow.</p>
          </div>

          <div className="grid md:grid-cols-3 gap-8 max-w-4xl mx-auto">
            {[
              {
                name: 'Starter',
                price: '$99',
                period: '/month',
                description: 'Perfect for small portfolios',
                features: ['Up to 100 tickets/month', '1 property', '5 vendors', 'Email notifications', 'Activity timeline'],
                cta: 'Get started',
                highlighted: false,
              },
              {
                name: 'Growth',
                price: '$299',
                period: '/month',
                description: 'For growing property managers',
                features: ['Unlimited tickets', 'Up to 10 properties', 'Unlimited vendors', 'AI categorization', 'AI vendor dispatch', 'SLA tracking', 'Priority support'],
                cta: 'Start free trial',
                highlighted: true,
              },
              {
                name: 'Pro',
                price: '$599',
                period: '/month',
                description: 'For large portfolios',
                features: ['Everything in Growth', 'Unlimited properties', 'Advanced SLA reporting', 'White-label option', 'Dedicated CSM', 'API access'],
                cta: 'Get started',
                highlighted: false,
              },
            ].map(plan => (
              <div
                key={plan.name}
                className={`rounded-xl p-8 border ${
                  plan.highlighted
                    ? 'border-indigo-300 bg-indigo-50 ring-2 ring-indigo-200'
                    : 'border-gray-200 bg-white'
                }`}
              >
                {plan.highlighted && (
                  <div className="text-xs font-medium text-indigo-600 bg-indigo-100 px-2 py-1 rounded-full w-fit mb-4">
                    Most popular
                  </div>
                )}
                <h3 className="text-lg font-semibold text-gray-900 mb-1">{plan.name}</h3>
                <p className="text-sm text-gray-500 mb-4">{plan.description}</p>
                <div className="flex items-baseline gap-1 mb-6">
                  <span className="text-4xl font-bold text-gray-900">{plan.price}</span>
                  <span className="text-gray-500">{plan.period}</span>
                </div>
                <ul className="space-y-2.5 mb-8">
                  {plan.features.map(f => (
                    <li key={f} className="flex items-center gap-2 text-sm text-gray-600">
                      <svg className="w-4 h-4 text-green-500 shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
                      </svg>
                      {f}
                    </li>
                  ))}
                </ul>
                <Link
                  href="/signup"
                  className={`block text-center py-2.5 rounded-lg text-sm font-medium transition-colors ${
                    plan.highlighted
                      ? 'bg-indigo-600 text-white hover:bg-indigo-700'
                      : 'bg-gray-900 text-white hover:bg-gray-700'
                  }`}
                >
                  {plan.cta}
                </Link>
              </div>
            ))}
          </div>
        </div>
      </section>

      {/* CTA */}
      <section className="bg-gray-900 py-20">
        <div className="max-w-3xl mx-auto px-6 text-center">
          <h2 className="text-3xl font-semibold text-white mb-4">
            Ready to eliminate maintenance chaos?
          </h2>
          <p className="text-gray-400 text-lg mb-8">
            Join property managers who have cut response times in half.
          </p>
          <Link
            href="/signup"
            className="inline-flex items-center gap-2 bg-white text-gray-900 px-8 py-3 rounded-lg font-medium hover:bg-gray-100 transition-colors"
          >
            Get started free
            <ArrowRight className="w-4 h-4" />
          </Link>
        </div>
      </section>

      {/* Footer */}
      <footer className="border-t border-gray-100 py-8">
        <div className="max-w-6xl mx-auto px-6 flex flex-col sm:flex-row items-center justify-between gap-4">
          <span className="text-sm font-medium text-gray-900">Maintena</span>
          <p className="text-sm text-gray-400">
            &copy; {new Date().getFullYear()} Maintena. The AI operations layer for property maintenance.
          </p>
        </div>
      </footer>
    </div>
  )
}

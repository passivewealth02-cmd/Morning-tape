import Link from 'next/link'
import { ArrowRight, Check, ChevronRight } from 'lucide-react'
import type { LandingPage } from '@/lib/landing-pages'
import { getLandingPage } from '@/lib/landing-pages'
import { SiteHeader } from './site-header'
import { SiteFooter } from './site-footer'

export function LandingTemplate({ page }: { page: LandingPage }) {
  const related = page.related
    .map(getLandingPage)
    .filter((p): p is LandingPage => Boolean(p))

  return (
    <div className="min-h-screen bg-white">
      <SiteHeader />

      <main>
        {/* Breadcrumb */}
        <nav aria-label="Breadcrumb" className="max-w-4xl mx-auto px-6 pt-8">
          <ol className="flex items-center gap-1.5 text-xs text-gray-400">
            <li><Link href="/" className="hover:text-gray-700">Home</Link></li>
            <li aria-hidden><ChevronRight className="w-3 h-3" /></li>
            <li className="text-gray-600">{page.eyebrow}</li>
          </ol>
        </nav>

        {/* Hero */}
        <section className="max-w-4xl mx-auto px-6 pt-6 pb-12">
          <p className="text-xs font-semibold uppercase tracking-wide text-indigo-600 mb-3">{page.eyebrow}</p>
          <h1 className="text-4xl md:text-5xl font-semibold tracking-tight text-gray-900 leading-tight mb-5">
            {page.h1}
          </h1>
          <p className="text-lg text-gray-500 leading-relaxed max-w-3xl mb-8">{page.intro}</p>
          <div className="flex flex-col sm:flex-row gap-3">
            <Link
              href="/signup"
              className="inline-flex items-center justify-center gap-2 bg-gray-900 text-white px-6 py-3 rounded-lg font-medium hover:bg-gray-700 transition-colors"
            >
              {page.ctaLabel}
              <ArrowRight className="w-4 h-4" />
            </Link>
            <Link
              href="/#features"
              className="inline-flex items-center justify-center gap-2 bg-gray-50 text-gray-700 px-6 py-3 rounded-lg font-medium hover:bg-gray-100 transition-colors border border-gray-200"
            >
              Explore all features
            </Link>
          </div>
        </section>

        {/* Benefits */}
        <section className="bg-gray-50 py-14 border-y border-gray-100">
          <div className="max-w-4xl mx-auto px-6">
            <div className="grid sm:grid-cols-2 lg:grid-cols-4 gap-6">
              {page.benefits.map(b => (
                <div key={b.title}>
                  <h2 className="text-sm font-semibold text-gray-900 mb-1.5">{b.title}</h2>
                  <p className="text-sm text-gray-500 leading-relaxed">{b.body}</p>
                </div>
              ))}
            </div>
          </div>
        </section>

        {/* Content sections */}
        <div className="max-w-3xl mx-auto px-6 py-16 space-y-14">
          {page.sections.map(s => (
            <section key={s.heading}>
              <h2 className="text-2xl font-semibold text-gray-900 tracking-tight mb-4">{s.heading}</h2>
              {s.body.map((p, i) => (
                <p key={i} className="text-base text-gray-600 leading-relaxed mb-4">{p}</p>
              ))}
              {s.bullets && (
                <ul className="space-y-2.5 mt-4">
                  {s.bullets.map(b => (
                    <li key={b} className="flex items-start gap-2.5 text-gray-600">
                      <Check className="w-5 h-5 text-indigo-600 shrink-0 mt-0.5" />
                      <span>{b}</span>
                    </li>
                  ))}
                </ul>
              )}
            </section>
          ))}
        </div>

        {/* FAQ */}
        <section id="faq" className="bg-gray-50 py-16 border-y border-gray-100">
          <div className="max-w-3xl mx-auto px-6">
            <h2 className="text-2xl font-semibold text-gray-900 tracking-tight mb-8">Frequently asked questions</h2>
            <div className="space-y-3">
              {page.faqs.map(f => (
                <details key={f.q} className="group rounded-xl border border-gray-200 bg-white p-5">
                  <summary className="flex items-center justify-between cursor-pointer list-none text-base font-medium text-gray-900">
                    {f.q}
                    <ChevronRight className="w-4 h-4 text-gray-400 transition-transform group-open:rotate-90" />
                  </summary>
                  <p className="mt-3 text-sm text-gray-600 leading-relaxed">{f.a}</p>
                </details>
              ))}
            </div>
          </div>
        </section>

        {/* Related solutions — internal linking */}
        {related.length > 0 && (
          <section className="max-w-4xl mx-auto px-6 py-16">
            <h2 className="text-xl font-semibold text-gray-900 mb-6">Explore related solutions</h2>
            <div className="grid sm:grid-cols-3 gap-4">
              {related.map(r => (
                <Link
                  key={r.slug}
                  href={`/${r.slug}`}
                  className="group rounded-xl border border-gray-200 p-5 hover:border-indigo-300 hover:bg-indigo-50/40 transition-colors"
                >
                  <h3 className="text-sm font-semibold text-gray-900 mb-1">{r.eyebrow}</h3>
                  <span className="inline-flex items-center gap-1 text-sm text-indigo-600">
                    Learn more <ArrowRight className="w-3.5 h-3.5 group-hover:translate-x-0.5 transition-transform" />
                  </span>
                </Link>
              ))}
            </div>
          </section>
        )}

        {/* CTA */}
        <section className="bg-gray-900 py-16">
          <div className="max-w-3xl mx-auto px-6 text-center">
            <h2 className="text-3xl font-semibold text-white mb-4">Ready to streamline property maintenance?</h2>
            <p className="text-gray-400 text-lg mb-8">14-day free trial. Cancel before day 14 and pay nothing.</p>
            <Link
              href="/signup"
              className="inline-flex items-center gap-2 bg-white text-gray-900 px-8 py-3 rounded-lg font-medium hover:bg-gray-100 transition-colors"
            >
              Start free trial
              <ArrowRight className="w-4 h-4" />
            </Link>
          </div>
        </section>
      </main>

      <SiteFooter />
    </div>
  )
}

import Link from 'next/link'

const solutions = [
  { href: '/property-maintenance-software', label: 'Property maintenance software' },
  { href: '/vendor-dispatch-software', label: 'Vendor dispatch software' },
  { href: '/maintenance-ticket-management', label: 'Maintenance ticket management' },
  { href: '/tenant-maintenance-requests', label: 'Tenant maintenance requests' },
  { href: '/ai-property-management-tools', label: 'AI property management tools' },
]

const product = [
  { href: '/#features', label: 'Features' },
  { href: '/#pricing', label: 'Pricing' },
  { href: '/#faq', label: 'FAQ' },
  { href: '/signup', label: 'Get started' },
]

export function SiteFooter() {
  return (
    <footer className="border-t border-gray-100 bg-white">
      <div className="max-w-6xl mx-auto px-6 py-12">
        <div className="grid gap-8 sm:grid-cols-2 md:grid-cols-4">
          <div>
            <span className="text-sm font-semibold text-gray-900">Maintena</span>
            <p className="mt-3 text-sm text-gray-500 leading-relaxed max-w-xs">
              AI-powered property maintenance software that captures requests, dispatches vendors, and tracks every repair.
            </p>
          </div>
          <nav aria-label="Solutions">
            <h2 className="text-xs font-semibold uppercase tracking-wide text-gray-400 mb-3">Solutions</h2>
            <ul className="space-y-2">
              {solutions.map(l => (
                <li key={l.href}>
                  <Link href={l.href} className="text-sm text-gray-600 hover:text-gray-900 transition-colors">{l.label}</Link>
                </li>
              ))}
            </ul>
          </nav>
          <nav aria-label="Product">
            <h2 className="text-xs font-semibold uppercase tracking-wide text-gray-400 mb-3">Product</h2>
            <ul className="space-y-2">
              {product.map(l => (
                <li key={l.href}>
                  <Link href={l.href} className="text-sm text-gray-600 hover:text-gray-900 transition-colors">{l.label}</Link>
                </li>
              ))}
            </ul>
          </nav>
          <nav aria-label="Account">
            <h2 className="text-xs font-semibold uppercase tracking-wide text-gray-400 mb-3">Account</h2>
            <ul className="space-y-2">
              <li><Link href="/login" className="text-sm text-gray-600 hover:text-gray-900 transition-colors">Sign in</Link></li>
              <li><Link href="/signup" className="text-sm text-gray-600 hover:text-gray-900 transition-colors">Sign up</Link></li>
            </ul>
          </nav>
        </div>
        <div className="mt-10 pt-6 border-t border-gray-100 flex flex-col sm:flex-row items-center justify-between gap-3">
          <p className="text-sm text-gray-400">
            &copy; {new Date().getFullYear()} Maintena. AI property maintenance software for property managers.
          </p>
        </div>
      </div>
    </footer>
  )
}

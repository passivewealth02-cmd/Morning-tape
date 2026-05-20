import Link from 'next/link'

export function SiteHeader() {
  return (
    <header className="border-b border-gray-100 bg-white/80 backdrop-blur-sm sticky top-0 z-50">
      <nav className="max-w-6xl mx-auto px-6 h-14 flex items-center justify-between" aria-label="Primary">
        <Link href="/" className="text-lg font-semibold tracking-tight text-gray-900">
          Maintena
        </Link>
        <div className="flex items-center gap-5">
          <Link href="/#features" className="hidden sm:inline text-sm text-gray-600 hover:text-gray-900 transition-colors">Features</Link>
          <Link href="/property-maintenance-software" className="hidden md:inline text-sm text-gray-600 hover:text-gray-900 transition-colors">Solutions</Link>
          <Link href="/#pricing" className="hidden sm:inline text-sm text-gray-600 hover:text-gray-900 transition-colors">Pricing</Link>
          <Link href="/#faq" className="hidden md:inline text-sm text-gray-600 hover:text-gray-900 transition-colors">FAQ</Link>
          <Link href="/login" className="text-sm text-gray-600 hover:text-gray-900 transition-colors">Sign in</Link>
          <Link href="/signup" className="text-sm bg-gray-900 text-white px-4 py-1.5 rounded-md hover:bg-gray-700 transition-colors">Sign up</Link>
        </div>
      </nav>
    </header>
  )
}

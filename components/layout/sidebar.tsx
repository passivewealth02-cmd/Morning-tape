'use client'

import { useState } from 'react'
import Link from 'next/link'
import { usePathname } from 'next/navigation'
import {
  LayoutDashboard,
  ClipboardList,
  Users,
  Building2,
  Settings,
  LogOut,
  Sparkles,
  Menu,
  X,
} from 'lucide-react'
import { MaintenaLogo } from '@/components/brand/logo'

const nav = [
  { href: '/dashboard', label: 'Dashboard', icon: LayoutDashboard },
  { href: '/dashboard/tickets', label: 'Tickets', icon: ClipboardList },
  { href: '/dashboard/vendors', label: 'Vendors', icon: Users },
  { href: '/dashboard/properties', label: 'Properties', icon: Building2 },
  { href: '/dashboard/settings#plan', label: 'Plan', icon: Sparkles },
]

function NavContents({ pathname, onNavigate }: { pathname: string; onNavigate?: () => void }) {
  return (
    <>
      <nav className="flex-1 px-3 py-4 space-y-0.5 overflow-y-auto">
        {nav.map(item => {
          const href = item.href.split('#')[0]
          const active = pathname === href || (href !== '/dashboard' && pathname.startsWith(href) && item.label !== 'Plan')
          return (
            <Link
              key={item.href}
              href={item.href}
              onClick={onNavigate}
              className={`flex items-center gap-2.5 px-3 py-2 rounded-md text-sm transition-colors ${
                active
                  ? 'bg-indigo-50 text-indigo-700 font-medium'
                  : 'text-gray-600 hover:bg-gray-50 hover:text-gray-900'
              }`}
            >
              <item.icon className="w-4 h-4 shrink-0" />
              {item.label}
            </Link>
          )
        })}
      </nav>

      <div className="px-3 py-4 border-t border-gray-100 space-y-0.5">
        <Link
          href="/dashboard/settings"
          onClick={onNavigate}
          className="flex items-center gap-2.5 px-3 py-2 rounded-md text-sm text-gray-600 hover:bg-gray-50 hover:text-gray-900 transition-colors"
        >
          <Settings className="w-4 h-4" />
          Settings
        </Link>
        <form action="/api/auth/logout" method="POST">
          <button
            type="submit"
            className="w-full flex items-center gap-2.5 px-3 py-2 rounded-md text-sm text-gray-600 hover:bg-gray-50 hover:text-gray-900 transition-colors"
          >
            <LogOut className="w-4 h-4" />
            Sign out
          </button>
        </form>
      </div>
    </>
  )
}

export function Sidebar() {
  const pathname = usePathname()
  const [open, setOpen] = useState(false)

  return (
    <>
      {/* Mobile top bar */}
      <div className="lg:hidden flex items-center justify-between px-4 h-14 border-b border-gray-100 bg-white sticky top-0 z-30 print:hidden">
        <MaintenaLogo />
        <button
          type="button"
          onClick={() => setOpen(true)}
          aria-label="Open menu"
          className="-mr-2 p-2 text-gray-600 hover:text-gray-900"
        >
          <Menu className="w-6 h-6" />
        </button>
      </div>

      {/* Mobile drawer */}
      {open && (
        <div className="lg:hidden fixed inset-0 z-50 print:hidden">
          <div className="absolute inset-0 bg-black/40" onClick={() => setOpen(false)} />
          <aside className="absolute left-0 top-0 bottom-0 w-64 max-w-[82%] bg-white flex flex-col shadow-xl">
            <div className="px-5 py-4 border-b border-gray-100 flex items-center justify-between">
              <MaintenaLogo />
              <button
                type="button"
                onClick={() => setOpen(false)}
                aria-label="Close menu"
                className="p-1 text-gray-500 hover:text-gray-800"
              >
                <X className="w-5 h-5" />
              </button>
            </div>
            <NavContents pathname={pathname} onNavigate={() => setOpen(false)} />
          </aside>
        </div>
      )}

      {/* Desktop sidebar */}
      <aside className="hidden lg:flex w-56 shrink-0 border-r border-gray-100 bg-white flex-col h-screen sticky top-0 print:hidden">
        <div className="px-5 py-4 border-b border-gray-100">
          <MaintenaLogo />
        </div>
        <NavContents pathname={pathname} />
      </aside>
    </>
  )
}

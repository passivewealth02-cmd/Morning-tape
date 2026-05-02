import type { Metadata, Viewport } from 'next'
import { Fraunces, Crimson_Pro, JetBrains_Mono } from 'next/font/google'
import { Analytics } from '@vercel/analytics/next'
import './globals.css'

const fraunces = Fraunces({ 
  subsets: ['latin'],
  variable: '--font-fraunces',
  display: 'swap',
})

const crimsonPro = Crimson_Pro({ 
  subsets: ['latin'],
  variable: '--font-crimson-pro',
  display: 'swap',
})

const jetbrainsMono = JetBrains_Mono({ 
  subsets: ['latin'],
  variable: '--font-jetbrains-mono',
  display: 'swap',
})

export const metadata: Metadata = {
  title: 'The Morning Tape | Daily AI Market Briefing',
  description: 'Your daily AI-generated market intelligence briefing. Concise analysis of market movements, sector performance, and economic events delivered fresh each trading day.',
  generator: 'v0.app',
  keywords: ['market briefing', 'AI analysis', 'stock market', 'financial news', 'trading intelligence'],
  authors: [{ name: 'The Morning Tape' }],
  icons: {
    icon: [
      {
        url: '/icon-light-32x32.png',
        media: '(prefers-color-scheme: light)',
      },
      {
        url: '/icon-dark-32x32.png',
        media: '(prefers-color-scheme: dark)',
      },
      {
        url: '/icon.svg',
        type: 'image/svg+xml',
      },
    ],
    apple: '/apple-icon.png',
  },
}

export const viewport: Viewport = {
  themeColor: '#F4EDE0',
  width: 'device-width',
  initialScale: 1,
}

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode
}>) {
  return (
    <html lang="en" className={`${fraunces.variable} ${crimsonPro.variable} ${jetbrainsMono.variable} bg-background`}>
      <body className="font-sans antialiased min-h-screen">
        {children}
        {process.env.NODE_ENV === 'production' && <Analytics />}
      </body>
    </html>
  )
}

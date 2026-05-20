import type { Metadata, Viewport } from 'next'
import { Inter, JetBrains_Mono } from 'next/font/google'
import { Analytics } from '@vercel/analytics/next'
import { SITE, organizationSchema, websiteSchema } from '@/lib/seo'
import { JsonLd } from '@/components/seo/json-ld'
import './globals.css'

const inter = Inter({
  subsets: ['latin'],
  variable: '--font-inter',
  display: 'swap',
})

const jetbrainsMono = JetBrains_Mono({
  subsets: ['latin'],
  variable: '--font-mono',
  display: 'swap',
})

export const metadata: Metadata = {
  metadataBase: new URL(SITE.url),
  title: {
    default: 'Maintena | AI Property Maintenance Software',
    template: '%s | Maintena',
  },
  description: SITE.description,
  applicationName: SITE.name,
  keywords: [
    'property maintenance software',
    'maintenance management software',
    'AI property maintenance software',
    'vendor dispatch software',
    'maintenance coordination platform',
    'tenant maintenance requests',
    'property management maintenance platform',
  ],
  authors: [{ name: 'Maintena' }],
  creator: 'Maintena',
  publisher: 'Maintena',
  alternates: { canonical: '/' },
  robots: {
    index: true,
    follow: true,
    googleBot: { index: true, follow: true, 'max-image-preview': 'large', 'max-snippet': -1 },
  },
  openGraph: {
    type: 'website',
    siteName: SITE.name,
    title: 'Maintena | AI Property Maintenance Software',
    description: SITE.description,
    url: SITE.url,
    locale: SITE.locale,
  },
  twitter: {
    card: 'summary_large_image',
    title: 'Maintena | AI Property Maintenance Software',
    description: SITE.description,
  },
}

export const viewport: Viewport = {
  width: 'device-width',
  initialScale: 1,
  themeColor: '#4f46e5',
}

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode
}>) {
  return (
    <html lang="en" className={`${inter.variable} ${jetbrainsMono.variable}`}>
      <body className="antialiased min-h-screen bg-background text-foreground">
        <JsonLd data={[organizationSchema(), websiteSchema()]} />
        {children}
        {process.env.NODE_ENV === 'production' && <Analytics />}
      </body>
    </html>
  )
}

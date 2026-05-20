import type { Metadata } from 'next'

export const SITE = {
  name: 'Maintena',
  legalName: 'Maintena',
  url: (process.env.NEXT_PUBLIC_APP_URL || 'https://trymaintena.com').replace(/\/$/, ''),
  tagline: 'AI Property Maintenance Software',
  description:
    'Maintena is AI-powered property maintenance software that captures tenant requests, triages them automatically, dispatches the right vendor, and tracks every repair to completion.',
  logo: '/icon.svg',
  ogImage: '/opengraph-image',
  locale: 'en_US',
}

/** Build page metadata with sensible SEO defaults (canonical, OG, Twitter). */
export function pageMetadata({
  title,
  description,
  path = '/',
  keywords,
  noindex = false,
}: {
  title: string
  description: string
  path?: string
  keywords?: string[]
  noindex?: boolean
}): Metadata {
  const canonical = path === '/' ? '/' : path.replace(/\/$/, '')
  const url = `${SITE.url}${canonical === '/' ? '' : canonical}`
  return {
    title: { absolute: title },
    description,
    keywords,
    alternates: { canonical },
    robots: noindex
      ? { index: false, follow: false }
      : { index: true, follow: true, 'max-image-preview': 'large', 'max-snippet': -1, 'max-video-preview': -1 },
    openGraph: {
      type: 'website',
      siteName: SITE.name,
      title,
      description,
      url,
      locale: SITE.locale,
    },
    twitter: {
      card: 'summary_large_image',
      title,
      description,
    },
  }
}

/* ---------------- JSON-LD schema builders ---------------- */

export function organizationSchema() {
  return {
    '@context': 'https://schema.org',
    '@type': 'Organization',
    name: SITE.name,
    legalName: SITE.legalName,
    url: SITE.url,
    logo: `${SITE.url}${SITE.logo}`,
    description: SITE.description,
    foundingDate: '2024',
    sameAs: ['https://x.com/trymaintena'],
    contactPoint: {
      '@type': 'ContactPoint',
      contactType: 'customer support',
      email: 'support@trymaintena.com',
      availableLanguage: ['English'],
    },
  }
}

export function websiteSchema() {
  return {
    '@context': 'https://schema.org',
    '@type': 'WebSite',
    name: SITE.name,
    url: SITE.url,
    description: SITE.description,
    publisher: { '@type': 'Organization', name: SITE.name, url: SITE.url },
  }
}

export function softwareApplicationSchema(opts?: { name?: string; description?: string }) {
  return {
    '@context': 'https://schema.org',
    '@type': 'SoftwareApplication',
    name: opts?.name || `${SITE.name} — ${SITE.tagline}`,
    applicationCategory: 'BusinessApplication',
    applicationSubCategory: 'Property Maintenance Software',
    operatingSystem: 'Web',
    description: opts?.description || SITE.description,
    url: SITE.url,
    offers: [
      { '@type': 'Offer', name: 'Starter', price: '99', priceCurrency: 'USD' },
      { '@type': 'Offer', name: 'Growth', price: '299', priceCurrency: 'USD' },
      { '@type': 'Offer', name: 'Pro', price: '599', priceCurrency: 'USD' },
    ],
    featureList: [
      'AI maintenance request triage',
      'Automated vendor dispatch',
      'Maintenance ticket tracking',
      'Tenant maintenance requests',
      'SLA tracking and alerts',
      'Audit timeline',
    ],
  }
}

export function faqSchema(faqs: { q: string; a: string }[]) {
  return {
    '@context': 'https://schema.org',
    '@type': 'FAQPage',
    mainEntity: faqs.map(f => ({
      '@type': 'Question',
      name: f.q,
      acceptedAnswer: { '@type': 'Answer', text: f.a },
    })),
  }
}

export function breadcrumbSchema(items: { name: string; path: string }[]) {
  return {
    '@context': 'https://schema.org',
    '@type': 'BreadcrumbList',
    itemListElement: items.map((item, i) => ({
      '@type': 'ListItem',
      position: i + 1,
      name: item.name,
      item: `${SITE.url}${item.path === '/' ? '' : item.path}`,
    })),
  }
}

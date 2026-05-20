import type { MetadataRoute } from 'next'
import { SITE } from '@/lib/seo'
import { LANDING_PAGES } from '@/lib/landing-pages'

export default function sitemap(): MetadataRoute.Sitemap {
  const now = new Date()
  const home: MetadataRoute.Sitemap[number] = {
    url: SITE.url,
    lastModified: now,
    changeFrequency: 'weekly',
    priority: 1,
  }
  const landing = LANDING_PAGES.map(p => ({
    url: `${SITE.url}/${p.slug}`,
    lastModified: now,
    changeFrequency: 'monthly' as const,
    priority: 0.8,
  }))
  return [home, ...landing]
}

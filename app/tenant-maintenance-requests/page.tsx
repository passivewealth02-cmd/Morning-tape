import { LandingTemplate } from '@/components/marketing/landing-template'
import { JsonLd } from '@/components/seo/json-ld'
import { getLandingPage } from '@/lib/landing-pages'
import { pageMetadata, faqSchema, breadcrumbSchema, softwareApplicationSchema } from '@/lib/seo'

const SLUG = 'tenant-maintenance-requests'
const page = getLandingPage(SLUG)!

export const metadata = pageMetadata({
  title: page.metaTitle,
  description: page.metaDescription,
  path: `/${SLUG}`,
  keywords: page.keywords,
})

export default function Page() {
  return (
    <>
      <JsonLd
        data={[
          breadcrumbSchema([{ name: 'Home', path: '/' }, { name: page.eyebrow, path: `/${SLUG}` }]),
          faqSchema(page.faqs),
          softwareApplicationSchema({ name: page.metaTitle, description: page.metaDescription }),
        ]}
      />
      <LandingTemplate page={page} />
    </>
  )
}

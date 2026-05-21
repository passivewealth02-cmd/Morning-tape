import { sql, type Organization } from '@/lib/db'
import { TenantSubmitForm } from '@/components/submit/tenant-submit-form'
import { MaintenaLogo } from '@/components/brand/logo'

interface Props {
  params: Promise<{ token: string }>
}

export const metadata = {
  title: 'Submit a maintenance request',
  robots: { index: false, follow: false },
}

export default async function TenantSubmitPage({ params }: Props) {
  const { token } = await params

  const orgs = (await sql`
    SELECT id, name FROM organizations WHERE inbox_token = ${token} LIMIT 1
  `) as unknown as Pick<Organization, 'id' | 'name'>[]

  if (orgs.length === 0) {
    return (
      <div className="min-h-screen flex items-center justify-center px-6 bg-gray-50">
        <div className="max-w-sm w-full text-center">
          <h1 className="text-lg font-semibold text-gray-900">Link not found</h1>
          <p className="text-sm text-gray-500 mt-2">
            This maintenance request link is invalid or has been disabled. Please contact your property manager.
          </p>
        </div>
      </div>
    )
  }

  const org = orgs[0]

  return (
    <div className="min-h-screen bg-gray-50 py-10 px-6">
      <div className="max-w-lg mx-auto">
        <div className="flex items-center justify-center mb-6">
          <MaintenaLogo />
        </div>
        <div className="text-center mb-6">
          <h1 className="text-2xl font-semibold text-gray-900">Submit a maintenance request</h1>
          <p className="text-gray-500 text-sm mt-2">
            For <span className="font-medium text-gray-700">{org.name}</span>. Tell us what needs fixing and we&apos;ll get it routed to the right person.
          </p>
        </div>
        <TenantSubmitForm token={token} />
      </div>
    </div>
  )
}

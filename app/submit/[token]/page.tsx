import { sql, type Organization } from '@/lib/db'
import { TenantSubmitForm } from '@/components/submit/tenant-submit-form'
import { MaintenaLogo } from '@/components/brand/logo'

interface Props {
  params: Promise<{ token: string }>
  searchParams: Promise<{ property?: string; unit?: string }>
}

export const metadata = {
  title: 'Submit a maintenance request',
  robots: { index: false, follow: false },
}

type PropertyOption = { id: string; name: string; address: string }
type UnitOption = { id: string; property_id: string; unit_number: string }

export default async function TenantSubmitPage({ params, searchParams }: Props) {
  const { token } = await params
  const { property: preProperty, unit: preUnit } = await searchParams

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

  const properties = (await sql`
    SELECT id, name, address FROM properties WHERE organization_id = ${org.id} ORDER BY name ASC
  `) as unknown as PropertyOption[]

  const units = (await sql`
    SELECT u.id, u.property_id, u.unit_number
    FROM units u
    JOIN properties p ON p.id = u.property_id
    WHERE p.organization_id = ${org.id}
    ORDER BY u.unit_number ASC
  `) as unknown as UnitOption[]

  // Validate any QR-provided pre-selection against this org's data.
  const presetProperty = properties.find(p => p.id === preProperty)?.id ?? ''
  const presetUnit =
    presetProperty && units.find(u => u.id === preUnit && u.property_id === presetProperty)?.id
      ? preUnit!
      : ''

  return (
    <div className="min-h-screen bg-gray-50">
      <div className="max-w-lg mx-auto px-5 py-8 sm:py-10">
        <div className="flex items-center justify-center mb-6">
          <MaintenaLogo />
        </div>
        <div className="text-center mb-6">
          <h1 className="text-2xl font-semibold text-gray-900">Report a maintenance issue</h1>
          <p className="text-gray-500 text-sm mt-2">
            for <span className="font-medium text-gray-700">{org.name}</span>
          </p>
        </div>
        <TenantSubmitForm
          token={token}
          properties={properties}
          units={units}
          presetProperty={presetProperty}
          presetUnit={presetUnit}
        />
      </div>
    </div>
  )
}

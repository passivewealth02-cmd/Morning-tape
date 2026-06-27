import { getSession } from '@/lib/auth'
import { sql, type Property, type Unit, type Organization } from '@/lib/db'
import Link from 'next/link'
import { notFound } from 'next/navigation'
import { ArrowLeft } from 'lucide-react'
import { PrintButton } from '@/components/properties/print-button'

interface Props {
  params: Promise<{ id: string }>
}

function baseUrl(): string {
  return (process.env.NEXT_PUBLIC_APP_URL || 'https://trymaintena.com').replace(/\/$/, '')
}

function qrSrc(data: string, size = 220): string {
  return `https://api.qrserver.com/v1/create-qr-code/?size=${size}x${size}&margin=8&data=${encodeURIComponent(data)}`
}

export default async function PropertyQrPage({ params }: Props) {
  const session = await getSession()
  if (!session) return null
  const orgId = session.user.organization_id!
  const { id } = await params

  const orgRows = (await sql`SELECT inbox_token, name FROM organizations WHERE id = ${orgId}`) as unknown as Pick<
    Organization,
    'inbox_token' | 'name'
  >[]
  const token = orgRows[0]?.inbox_token
  const orgName = orgRows[0]?.name ?? ''

  const propRows = (await sql`
    SELECT * FROM properties WHERE id = ${id} AND organization_id = ${orgId}
  `) as Property[]
  const property = propRows[0]
  if (!property) notFound()

  const units = (await sql`
    SELECT * FROM units WHERE property_id = ${id} ORDER BY unit_number ASC
  `) as Unit[]

  const base = baseUrl()
  const propertyUrl = `${base}/submit/${token}?property=${property.id}`

  return (
    <div className="p-6">
      <div className="print:hidden">
        <Link href={`/dashboard/properties/${property.id}`} className="inline-flex items-center gap-1.5 text-sm text-gray-500 hover:text-gray-900 mb-4">
          <ArrowLeft className="w-4 h-4" />
          Back to {property.name}
        </Link>
        <div className="flex items-start justify-between gap-4 mb-6">
          <div>
            <h1 className="text-xl font-semibold text-gray-900">QR codes — {property.name}</h1>
            <p className="text-sm text-gray-500 mt-0.5 max-w-xl">
              Print these and post them on unit doors, fridges, or mailrooms. When a tenant scans one, the
              request form opens already knowing their exact property and unit — no typing, no mistakes.
            </p>
          </div>
          <PrintButton />
        </div>
        {!token && (
          <p className="mb-4 text-sm text-amber-700 bg-amber-50 border border-amber-100 rounded-lg px-4 py-2">
            Your intake link isn’t set up yet — finish onboarding so these QR codes work.
          </p>
        )}
      </div>

      {/* Print header (only shows on paper) */}
      <div className="hidden print:block mb-6">
        <h1 className="text-lg font-semibold">{orgName} — {property.name}</h1>
        <p className="text-sm text-gray-600">Scan to submit a maintenance request</p>
      </div>

      <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-4 gap-4 print:grid-cols-3">
        {/* General property QR */}
        <div className="rounded-xl border border-gray-200 bg-white p-4 text-center break-inside-avoid">
          <img src={qrSrc(propertyUrl)} alt={`QR for ${property.name}`} className="w-full max-w-[180px] mx-auto" />
          <p className="mt-2 text-sm font-semibold text-gray-900">{property.name}</p>
          <p className="text-xs text-gray-500">General / lobby</p>
        </div>

        {units.map(u => {
          const url = `${base}/submit/${token}?property=${property.id}&unit=${u.id}`
          return (
            <div key={u.id} className="rounded-xl border border-gray-200 bg-white p-4 text-center break-inside-avoid">
              <img src={qrSrc(url)} alt={`QR for unit ${u.unit_number}`} className="w-full max-w-[180px] mx-auto" />
              <p className="mt-2 text-sm font-semibold text-gray-900">Unit {u.unit_number}</p>
              <p className="text-xs text-gray-500">{property.name}</p>
            </div>
          )
        })}
      </div>

      {units.length === 0 && (
        <p className="mt-4 text-sm text-gray-500 print:hidden">
          No units yet — add units on the{' '}
          <Link href={`/dashboard/properties/${property.id}`} className="text-indigo-600 hover:underline">
            property page
          </Link>{' '}
          to generate a QR code for each one.
        </p>
      )}
    </div>
  )
}

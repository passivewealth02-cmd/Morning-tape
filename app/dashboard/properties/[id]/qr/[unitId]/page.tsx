import { getSession } from '@/lib/auth'
import { sql, type Property, type Unit, type Organization } from '@/lib/db'
import Link from 'next/link'
import { notFound } from 'next/navigation'
import { ArrowLeft } from 'lucide-react'
import { QrActions } from '@/components/properties/qr-actions'

interface Props {
  params: Promise<{ id: string; unitId: string }>
}

function baseUrl(): string {
  return (process.env.NEXT_PUBLIC_APP_URL || 'https://trymaintena.com').replace(/\/$/, '')
}

function qrSrc(data: string, size = 360): string {
  return `https://api.qrserver.com/v1/create-qr-code/?size=${size}x${size}&margin=12&data=${encodeURIComponent(data)}`
}

export default async function SingleUnitQrPage({ params }: Props) {
  const session = await getSession()
  if (!session) return null
  const orgId = session.user.organization_id!
  const { id, unitId } = await params

  const orgRows = (await sql`
    SELECT inbox_token, name FROM organizations WHERE id = ${orgId}
  `) as unknown as Pick<Organization, 'inbox_token' | 'name'>[]
  const token = orgRows[0]?.inbox_token
  const orgName = orgRows[0]?.name ?? ''

  const propRows = (await sql`
    SELECT * FROM properties WHERE id = ${id} AND organization_id = ${orgId}
  `) as Property[]
  const property = propRows[0]
  if (!property) notFound()

  const unitRows = (await sql`
    SELECT * FROM units WHERE id = ${unitId} AND property_id = ${id}
  `) as Unit[]
  const unit = unitRows[0]
  if (!unit) notFound()

  const submitUrl = `${baseUrl()}/submit/${token}?property=${property.id}&unit=${unit.id}`
  const img = qrSrc(submitUrl)
  const title = `Unit ${unit.unit_number} — ${property.name}`
  const fileName = `QR-${property.name}-Unit-${unit.unit_number}.png`.replace(/[^a-zA-Z0-9.-]+/g, '-')

  return (
    <div className="p-6">
      <div className="print:hidden">
        <Link
          href={`/dashboard/properties/${property.id}/qr`}
          className="inline-flex items-center gap-1.5 text-sm text-gray-500 hover:text-gray-900 mb-6"
        >
          <ArrowLeft className="w-4 h-4" />
          Back to all QR codes
        </Link>
      </div>

      <div className="max-w-md mx-auto">
        <div className="bg-white rounded-2xl border border-gray-200 p-8 text-center shadow-sm break-inside-avoid print:border-0 print:shadow-none">
          <div className="hidden print:block mb-3 text-sm text-gray-600">{orgName}</div>
          {/* eslint-disable-next-line @next/next/no-img-element */}
          <img src={img} alt={`QR code for ${title}`} className="w-full max-w-[320px] mx-auto" />
          <p className="mt-4 text-xl font-semibold text-gray-900">Unit {unit.unit_number}</p>
          <p className="text-sm text-gray-500">{property.name}</p>
          <p className="mt-3 text-xs font-medium text-gray-600 print:block">Scan to submit a maintenance request</p>
          <p className="mt-2 text-[11px] text-gray-400 break-all print:hidden">{submitUrl}</p>
        </div>

        <div className="mt-5">
          <QrActions imageUrl={img} downloadName={fileName} submitUrl={submitUrl} title={title} />
        </div>
      </div>
    </div>
  )
}

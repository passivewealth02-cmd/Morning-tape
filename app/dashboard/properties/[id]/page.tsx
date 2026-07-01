import { getSession } from '@/lib/auth'
import { sql } from '@/lib/db'
import type { Property, Unit } from '@/lib/db'
import Link from 'next/link'
import { notFound } from 'next/navigation'
import { ArrowLeft, MapPin, QrCode } from 'lucide-react'
import { UnitManager } from '@/components/properties/unit-manager'
import { DeleteButton } from '@/components/ui/delete-button'

export default async function PropertyDetailPage({
  params,
}: {
  params: Promise<{ id: string }>
}) {
  const session = await getSession()
  if (!session) return null
  const orgId = session.user.organization_id!
  const { id } = await params

  const rows = (await sql`
    SELECT * FROM properties WHERE id = ${id} AND organization_id = ${orgId}
  `) as Property[]
  const property = rows[0]
  if (!property) notFound()

  const units = (await sql`
    SELECT * FROM units WHERE property_id = ${id} ORDER BY unit_number ASC
  `) as Unit[]

  return (
    <div className="p-6">
      <Link href="/dashboard/properties" className="inline-flex items-center gap-1.5 text-sm text-gray-500 hover:text-gray-900 mb-4">
        <ArrowLeft className="w-4 h-4" />
        Back to properties
      </Link>

      <div className="mb-6 flex items-start justify-between gap-4">
        <div>
          <h1 className="text-xl font-semibold text-gray-900">{property.name}</h1>
          <p className="text-sm text-gray-500 mt-0.5 inline-flex items-center gap-1.5">
            <MapPin className="w-3.5 h-3.5 text-gray-400" />
            {property.address}
            {(property.city || property.province) && ` · ${[property.city, property.province].filter(Boolean).join(', ')}`}
          </p>
        </div>
        <div className="flex items-center gap-2 shrink-0">
          <Link
            href={`/dashboard/properties/${property.id}/qr`}
            className="inline-flex items-center gap-1.5 text-sm font-medium border border-gray-200 hover:border-indigo-200 hover:bg-indigo-50 text-gray-700 px-3 py-2 rounded-lg transition-colors"
          >
            <QrCode className="w-4 h-4" />
            QR codes
          </Link>
          <DeleteButton
            endpoint={`/api/properties/${property.id}`}
            redirectTo="/dashboard/properties"
            label="Delete"
            confirmLabel="Delete property"
          />
        </div>
      </div>

      <UnitManager propertyId={property.id} units={units} />
    </div>
  )
}

import { getSession } from '@/lib/auth'
import { sql } from '@/lib/db'
import type { Property } from '@/lib/db'
import Link from 'next/link'
import { Building2, DoorOpen, ChevronRight } from 'lucide-react'
import { PropertyAddPanel } from '@/components/properties/property-add-panel'

export default async function PropertiesPage() {
  const session = await getSession()
  if (!session) return null

  const orgId = session.user.organization_id!

  const properties = await sql`
    SELECT
      p.*,
      COUNT(DISTINCT u.id)::int AS unit_count,
      COUNT(DISTINCT t.id) FILTER (WHERE t.status NOT IN ('completed', 'cancelled'))::int AS open_tickets
    FROM properties p
    LEFT JOIN units u ON u.property_id = p.id
    LEFT JOIN maintenance_tickets t ON t.property_id = p.id
    WHERE p.organization_id = ${orgId}
    GROUP BY p.id
    ORDER BY p.name ASC
  ` as (Property & { unit_count: number; open_tickets: number })[]

  const totalUnits = properties.reduce((sum, p) => sum + p.unit_count, 0)

  return (
    <div className="p-6">
      <div className="flex items-center justify-between mb-6">
        <div>
          <h1 className="text-xl font-semibold text-gray-900">Properties</h1>
          <p className="text-sm text-gray-500 mt-0.5">
            {properties.length} propert{properties.length !== 1 ? 'ies' : 'y'} · {totalUnits} unit{totalUnits !== 1 ? 's' : ''}
          </p>
        </div>
      </div>

      <div className="grid lg:grid-cols-3 gap-6">
        {/* Property list */}
        <div className="lg:col-span-2">
          {properties.length === 0 ? (
            <div className="bg-white rounded-xl border border-dashed border-gray-200 p-12 text-center">
              <Building2 className="w-8 h-8 text-gray-300 mx-auto mb-3" />
              <p className="text-gray-500 mb-1">No properties yet</p>
              <p className="text-sm text-gray-400">Add a property — and generate all its units at once — to start tracking maintenance by location.</p>
            </div>
          ) : (
            <div className="space-y-3">
              {properties.map(prop => (
                <Link
                  key={prop.id}
                  href={`/dashboard/properties/${prop.id}`}
                  className="block bg-white rounded-xl border border-gray-200 p-5 hover:border-indigo-200 hover:shadow-sm transition-all"
                >
                  <div className="flex items-start justify-between">
                    <div className="min-w-0">
                      <h3 className="font-semibold text-gray-900">{prop.name}</h3>
                      <p className="text-sm text-gray-500 mt-0.5 truncate">{prop.address}</p>
                      {(prop.city || prop.province) && (
                        <p className="text-xs text-gray-400 mt-0.5">{[prop.city, prop.province].filter(Boolean).join(', ')}</p>
                      )}
                      <div className="flex items-center gap-3 mt-2">
                        <span className="inline-flex items-center gap-1 text-xs text-gray-500">
                          <DoorOpen className="w-3.5 h-3.5 text-gray-400" />
                          {prop.unit_count} unit{prop.unit_count !== 1 ? 's' : ''}
                        </span>
                        {prop.open_tickets > 0 && (
                          <span className="text-xs font-medium text-indigo-600">
                            {prop.open_tickets} open ticket{prop.open_tickets !== 1 ? 's' : ''}
                          </span>
                        )}
                      </div>
                    </div>
                    <ChevronRight className="w-4 h-4 text-gray-300 shrink-0 mt-1" />
                  </div>
                </Link>
              ))}
            </div>
          )}
        </div>

        {/* Add property panel (single + bulk) */}
        <div>
          <PropertyAddPanel />
        </div>
      </div>
    </div>
  )
}

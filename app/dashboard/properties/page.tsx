import { getSession } from '@/lib/auth'
import { sql } from '@/lib/db'
import type { Property } from '@/lib/db'
import Link from 'next/link'
import { Plus, Building2 } from 'lucide-react'
import { NewPropertyForm } from '@/components/properties/new-property-form'

export default async function PropertiesPage() {
  const session = await getSession()
  if (!session) return null

  const orgId = session.user.organization_id!

  const properties = await sql`
    SELECT
      p.*,
      COUNT(t.id) FILTER (WHERE t.status NOT IN ('completed', 'cancelled'))::int AS open_tickets
    FROM properties p
    LEFT JOIN maintenance_tickets t ON t.property_id = p.id
    WHERE p.organization_id = ${orgId}
    GROUP BY p.id
    ORDER BY p.name ASC
  ` as (Property & { open_tickets: number })[]

  return (
    <div className="p-6">
      <div className="flex items-center justify-between mb-6">
        <div>
          <h1 className="text-xl font-semibold text-gray-900">Properties</h1>
          <p className="text-sm text-gray-500 mt-0.5">{properties.length} propert{properties.length !== 1 ? 'ies' : 'y'}</p>
        </div>
      </div>

      <div className="grid lg:grid-cols-3 gap-6">
        {/* Property list */}
        <div className="lg:col-span-2">
          {properties.length === 0 ? (
            <div className="bg-white rounded-xl border border-dashed border-gray-200 p-12 text-center">
              <Building2 className="w-8 h-8 text-gray-300 mx-auto mb-3" />
              <p className="text-gray-500 mb-1">No properties yet</p>
              <p className="text-sm text-gray-400">Add a property to start tracking maintenance requests by location.</p>
            </div>
          ) : (
            <div className="space-y-3">
              {properties.map(prop => (
                <div key={prop.id} className="bg-white rounded-xl border border-gray-200 p-5">
                  <div className="flex items-start justify-between">
                    <div>
                      <h3 className="font-semibold text-gray-900">{prop.name}</h3>
                      <p className="text-sm text-gray-500 mt-0.5">{prop.address}</p>
                      {(prop.city || prop.province) && (
                        <p className="text-xs text-gray-400 mt-0.5">{[prop.city, prop.province].filter(Boolean).join(', ')}</p>
                      )}
                    </div>
                    <div className="text-right">
                      {prop.open_tickets > 0 ? (
                        <Link
                          href={`/dashboard/tickets?property_id=${prop.id}`}
                          className="text-sm font-medium text-indigo-600 hover:underline"
                        >
                          {prop.open_tickets} open ticket{prop.open_tickets !== 1 ? 's' : ''}
                        </Link>
                      ) : (
                        <span className="text-sm text-gray-400">No open tickets</span>
                      )}
                    </div>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* Add property form */}
        <div>
          <div className="bg-white rounded-xl border border-gray-200 p-5">
            <h2 className="text-sm font-semibold text-gray-700 mb-4 flex items-center gap-2">
              <Plus className="w-4 h-4" />
              Add Property
            </h2>
            <NewPropertyForm />
          </div>
        </div>
      </div>
    </div>
  )
}

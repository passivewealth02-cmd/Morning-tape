import { getSession } from '@/lib/auth'
import { sql } from '@/lib/db'
import type { Vendor } from '@/lib/db'
import Link from 'next/link'
import { Plus, Star, Phone, Mail } from 'lucide-react'
import { DeleteButton } from '@/components/ui/delete-button'

const AVAILABILITY_COLORS = {
  available: 'bg-green-100 text-green-700',
  busy: 'bg-yellow-100 text-yellow-700',
  unavailable: 'bg-gray-100 text-gray-500',
}

const INSURANCE_COLORS = {
  verified: 'bg-green-100 text-green-700',
  expired: 'bg-red-100 text-red-700',
  unknown: 'bg-gray-100 text-gray-500',
}

export default async function VendorsPage() {
  const session = await getSession()
  if (!session) return null

  const orgId = session.user.organization_id!

  const vendors = await sql`
    SELECT
      v.*,
      COUNT(t.id) FILTER (WHERE t.status NOT IN ('completed', 'cancelled'))::int AS active_tickets
    FROM vendors v
    LEFT JOIN maintenance_tickets t ON t.assigned_vendor_id = v.id
    WHERE v.organization_id = ${orgId}
    GROUP BY v.id
    ORDER BY v.name ASC
  ` as (Vendor & { active_tickets: number })[]

  return (
    <div className="p-6">
      <div className="flex items-center justify-between mb-6">
        <div>
          <h1 className="text-xl font-semibold text-gray-900">Vendors</h1>
          <p className="text-sm text-gray-500 mt-0.5">{vendors.length} vendor{vendors.length !== 1 ? 's' : ''}</p>
        </div>
        <Link
          href="/dashboard/vendors/new"
          className="inline-flex items-center gap-1.5 bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-medium px-4 py-2 rounded-lg transition-colors"
        >
          <Plus className="w-4 h-4" />
          Add Vendor
        </Link>
      </div>

      {vendors.length === 0 ? (
        <div className="bg-white rounded-xl border border-dashed border-gray-200 p-12 text-center">
          <p className="text-gray-500 mb-4">No vendors yet. Add your first vendor to start dispatching jobs.</p>
          <Link
            href="/dashboard/vendors/new"
            className="inline-flex items-center gap-1.5 bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-medium px-4 py-2 rounded-lg transition-colors"
          >
            <Plus className="w-4 h-4" />
            Add your first vendor
          </Link>
        </div>
      ) : (
        <div className="grid md:grid-cols-2 lg:grid-cols-3 gap-4">
          {vendors.map(vendor => (
            <div key={vendor.id} className="bg-white rounded-xl border border-gray-200 p-5 hover:shadow-sm transition-shadow">
              <div className="flex items-start justify-between mb-3">
                <div>
                  <h3 className="font-semibold text-gray-900">{vendor.name}</h3>
                  <p className="text-sm text-gray-500 capitalize mt-0.5">{vendor.trade_type}</p>
                </div>
                <span className={`text-xs px-2 py-0.5 rounded-full font-medium capitalize ${AVAILABILITY_COLORS[vendor.availability]}`}>
                  {vendor.availability}
                </span>
              </div>

              <div className="space-y-2 mb-4">
                {vendor.phone && (
                  <div className="flex items-center gap-2 text-sm text-gray-600">
                    <Phone className="w-3.5 h-3.5 text-gray-400" />
                    <a href={`tel:${vendor.phone}`} className="hover:text-indigo-600">{vendor.phone}</a>
                  </div>
                )}
                {vendor.email && (
                  <div className="flex items-center gap-2 text-sm text-gray-600">
                    <Mail className="w-3.5 h-3.5 text-gray-400" />
                    <a href={`mailto:${vendor.email}`} className="hover:text-indigo-600 truncate">{vendor.email}</a>
                  </div>
                )}
              </div>

              <div className="flex items-center justify-between pt-3 border-t border-gray-100">
                <div className="flex items-center gap-1">
                  <Star className="w-3.5 h-3.5 text-yellow-400 fill-yellow-400" />
                  <span className="text-sm font-medium text-gray-700">{vendor.rating || '—'}</span>
                </div>
                <div className="flex items-center gap-3">
                  <span className={`text-xs px-1.5 py-0.5 rounded font-medium ${INSURANCE_COLORS[vendor.insurance_status]}`}>
                    {vendor.insurance_status === 'verified' ? 'Insured' : vendor.insurance_status === 'expired' ? 'Expired' : 'Unknown'}
                  </span>
                  {vendor.active_tickets > 0 && (
                    <span className="text-xs text-gray-500">{vendor.active_tickets} active</span>
                  )}
                </div>
              </div>

              <div className="flex justify-end mt-3">
                <DeleteButton endpoint={`/api/vendors/${vendor.id}`} label="Delete" size="sm" confirmLabel="Delete vendor" />
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  )
}

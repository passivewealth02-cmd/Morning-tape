import { getSession } from '@/lib/auth'
import { sql } from '@/lib/db'
import { NewTicketForm } from '@/components/tickets/new-ticket-form'
import type { Property, Vendor } from '@/lib/db'

export default async function NewTicketPage() {
  const session = await getSession()
  if (!session) return null

  const orgId = session.user.organization_id!

  const [propertiesRaw, vendorsRaw] = await Promise.all([
    sql`SELECT id, name, address FROM properties WHERE organization_id = ${orgId} ORDER BY name`,
    sql`SELECT id, name, trade_type FROM vendors WHERE organization_id = ${orgId} AND availability != 'unavailable' ORDER BY name`,
  ])

  const properties = propertiesRaw as unknown as Property[]
  const vendors = vendorsRaw as unknown as Vendor[]

  return (
    <div className="p-6 max-w-2xl">
      <div className="mb-6">
        <h1 className="text-xl font-semibold text-gray-900">New Maintenance Request</h1>
        <p className="text-sm text-gray-500 mt-0.5">AI will automatically categorize and prioritize this ticket.</p>
      </div>
      <NewTicketForm properties={properties} vendors={vendors} />
    </div>
  )
}

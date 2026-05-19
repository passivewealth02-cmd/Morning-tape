import { getSession } from '@/lib/auth'
import { sql, type Organization } from '@/lib/db'

export default async function SettingsPage() {
  const session = await getSession()
  if (!session) return null

  const orgRows = (await sql`
    SELECT * FROM organizations WHERE id = ${session.user.organization_id}
  `) as unknown as Organization[]
  const org = orgRows[0]

  return (
    <div className="p-6 max-w-2xl">
      <div className="mb-6">
        <h1 className="text-xl font-semibold text-gray-900">Settings</h1>
        <p className="text-sm text-gray-500 mt-0.5">Manage your account and workspace.</p>
      </div>

      <section className="bg-white rounded-lg border border-gray-200 p-6 mb-4">
        <h2 className="text-sm font-semibold text-gray-900 mb-4">Your account</h2>
        <dl className="space-y-3 text-sm">
          <div className="flex justify-between">
            <dt className="text-gray-500">Email</dt>
            <dd className="text-gray-900">{session.user.email}</dd>
          </div>
          {session.user.name && (
            <div className="flex justify-between">
              <dt className="text-gray-500">Name</dt>
              <dd className="text-gray-900">{session.user.name}</dd>
            </div>
          )}
          <div className="flex justify-between">
            <dt className="text-gray-500">Role</dt>
            <dd className="text-gray-900 capitalize">{session.user.role}</dd>
          </div>
        </dl>
      </section>

      {org && (
        <section className="bg-white rounded-lg border border-gray-200 p-6">
          <h2 className="text-sm font-semibold text-gray-900 mb-4">Workspace</h2>
          <dl className="space-y-3 text-sm">
            <div className="flex justify-between">
              <dt className="text-gray-500">Organization</dt>
              <dd className="text-gray-900">{org.name}</dd>
            </div>
            <div className="flex justify-between">
              <dt className="text-gray-500">Slug</dt>
              <dd className="text-gray-900 font-mono text-xs">{org.slug}</dd>
            </div>
            <div className="flex justify-between">
              <dt className="text-gray-500">Created</dt>
              <dd className="text-gray-900">
                {new Date(org.created_at).toLocaleDateString()}
              </dd>
            </div>
          </dl>
        </section>
      )}
    </div>
  )
}

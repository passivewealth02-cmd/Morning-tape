import { redirect } from 'next/navigation'
import { getSession } from '@/lib/auth'
import { sql, type Organization } from '@/lib/db'
import { Sidebar } from '@/components/layout/sidebar'
import { TrialBanner } from '@/components/layout/trial-banner'
import { trialDaysLeft, isTrialActive } from '@/lib/plans'

export default async function DashboardLayout({
  children,
}: {
  children: React.ReactNode
}) {
  const session = await getSession()

  if (!session) {
    redirect('/login')
  }

  if (!session.user.organization_id) {
    redirect('/onboarding')
  }

  const orgRows = (await sql`
    SELECT * FROM organizations WHERE id = ${session.user.organization_id}
  `) as unknown as Organization[]
  const org = orgRows[0]

  const showTrialBanner = org && (
    isTrialActive(org) ||
    (org.plan === 'trial' && org.trial_ends_at && new Date(org.trial_ends_at) < new Date())
  )

  return (
    <div className="flex h-screen bg-gray-50 overflow-hidden">
      <Sidebar />
      <main className="flex-1 overflow-y-auto">
        {showTrialBanner && org && (
          <TrialBanner
            daysLeft={trialDaysLeft(org)}
            expired={!isTrialActive(org)}
          />
        )}
        {children}
      </main>
    </div>
  )
}

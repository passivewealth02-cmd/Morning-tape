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

  const trialActive = org && isTrialActive(org)
  const trialExpiredNoCard = org && org.plan === 'trial' && org.trial_ends_at && new Date(org.trial_ends_at) < new Date()
  const pastDue = org && org.plan_status === 'past_due'
  const showTrialBanner = trialActive || trialExpiredNoCard || pastDue

  return (
    <div className="flex flex-col lg:flex-row h-screen bg-gray-50 overflow-hidden print:block print:h-auto print:overflow-visible">
      <Sidebar />
      <main className="flex-1 min-w-0 overflow-y-auto print:overflow-visible">
        {showTrialBanner && org && (
          <div className="print:hidden">
            <TrialBanner
              daysLeft={trialDaysLeft(org)}
              expired={!trialActive && !pastDue}
              pastDue={!!pastDue}
            />
          </div>
        )}
        {children}
      </main>
    </div>
  )
}

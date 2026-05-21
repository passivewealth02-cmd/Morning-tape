import 'server-only'
import { Resend } from 'resend'

const resend = new Resend(process.env.RESEND_API_KEY)

// Escape user-supplied values before interpolating them into HTML email bodies
// to prevent HTML/phishing-link injection in notification emails.
function escapeHtml(value: string | null | undefined): string {
  if (value == null) return ''
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
}

export async function sendMagicLinkEmail(email: string, token: string, baseUrl?: string): Promise<boolean> {
  const resolvedBase = baseUrl || process.env.NEXT_PUBLIC_APP_URL || 'http://localhost:3000'
  const magicLink = `${resolvedBase}/verify?token=${token}`

  try {
    const { error } = await resend.emails.send({
      from: 'Maintena <onboarding@resend.dev>',
      to: email,
      subject: 'Sign in to Maintena',
      html: `
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"></head>
<body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background-color: #f9fafb; padding: 40px 20px; margin: 0;">
  <div style="max-width: 480px; margin: 0 auto; background-color: #ffffff; border-radius: 12px; border: 1px solid #e5e7eb; overflow: hidden;">
    <div style="padding: 24px; border-bottom: 1px solid #f3f4f6; display: flex; align-items: center; gap: 10px;">
      <div style="width: 28px; height: 28px; background-color: #4f46e5; border-radius: 6px; display: flex; align-items: center; justify-content: center;">
        <span style="color: white; font-size: 14px; font-weight: bold;">M</span>
      </div>
      <span style="font-size: 15px; font-weight: 600; color: #111827;">Maintena</span>
    </div>

    <div style="padding: 32px 24px;">
      <h2 style="font-size: 20px; font-weight: 600; color: #111827; margin: 0 0 8px 0;">Sign in to Maintena</h2>
      <p style="font-size: 14px; color: #6b7280; margin: 0 0 24px 0; line-height: 1.5;">
        Click the button below to sign in. This link expires in 15 minutes.
      </p>

      <a href="${magicLink}" style="display: inline-block; background-color: #4f46e5; color: #ffffff; padding: 12px 24px; border-radius: 8px; text-decoration: none; font-size: 14px; font-weight: 500;">
        Sign in to Maintena
      </a>

      <p style="font-size: 12px; color: #9ca3af; margin: 24px 0 0 0; line-height: 1.5;">
        If you didn't request this, you can safely ignore this email. Your account won't be affected.
      </p>
    </div>

    <div style="padding: 16px 24px; background-color: #f9fafb; border-top: 1px solid #f3f4f6;">
      <p style="font-size: 11px; color: #9ca3af; margin: 0;">
        &copy; ${new Date().getFullYear()} Maintena. The AI operations layer for property maintenance.
      </p>
    </div>
  </div>
</body>
</html>
      `,
    })

    if (error) {
      console.error('Failed to send magic link email:', error)
      return false
    }

    return true
  } catch (error) {
    console.error('Error sending email:', error)
    return false
  }
}

export async function sendVendorAssignmentEmail(
  vendorEmail: string,
  vendorName: string,
  ticketTitle: string,
  ticketId: string,
  propertyName?: string
): Promise<boolean> {
  const baseUrl = process.env.NEXT_PUBLIC_APP_URL || 'http://localhost:3000'

  try {
    const { error } = await resend.emails.send({
      from: 'Maintena <onboarding@resend.dev>',
      to: vendorEmail,
      subject: `New job assigned: ${ticketTitle}`,
      html: `
<!DOCTYPE html>
<html>
<body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background-color: #f9fafb; padding: 40px 20px; margin: 0;">
  <div style="max-width: 480px; margin: 0 auto; background-color: #ffffff; border-radius: 12px; border: 1px solid #e5e7eb; padding: 32px 24px;">
    <p style="font-size: 15px; font-weight: 600; color: #111827; margin: 0 0 4px 0;">Maintena</p>
    <h2 style="font-size: 18px; font-weight: 600; color: #111827; margin: 16px 0 8px 0;">New job assigned to you</h2>
    <p style="font-size: 14px; color: #6b7280; margin: 0 0 16px 0;">Hi ${escapeHtml(vendorName)},</p>
    <p style="font-size: 14px; color: #374151; margin: 0 0 8px 0;"><strong>Job:</strong> ${escapeHtml(ticketTitle)}</p>
    ${propertyName ? `<p style="font-size: 14px; color: #374151; margin: 0 0 24px 0;"><strong>Property:</strong> ${escapeHtml(propertyName)}</p>` : ''}
    <p style="font-size: 12px; color: #9ca3af; margin: 24px 0 0 0;">© ${new Date().getFullYear()} Maintena</p>
  </div>
</body>
</html>
      `,
    })

    return !error
  } catch {
    return false
  }
}

export async function sendTenantVendorAssignedEmail(
  tenantEmail: string,
  tenantName: string | null,
  ticketTitle: string,
  vendorName: string,
  propertyName?: string | null
): Promise<boolean> {
  const greeting = tenantName ? `Hi ${escapeHtml(tenantName)},` : 'Hello,'
  try {
    const { error } = await resend.emails.send({
      from: 'Maintena <onboarding@resend.dev>',
      to: tenantEmail,
      subject: `Update on your maintenance request: ${ticketTitle}`,
      html: `
<!DOCTYPE html>
<html>
<body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background-color: #f9fafb; padding: 40px 20px; margin: 0;">
  <div style="max-width: 480px; margin: 0 auto; background-color: #ffffff; border-radius: 12px; border: 1px solid #e5e7eb; overflow: hidden;">
    <div style="padding: 20px 24px; border-bottom: 1px solid #f3f4f6;">
      <div style="display: inline-flex; align-items: center; gap: 8px;">
        <div style="width: 24px; height: 24px; background-color: #4f46e5; border-radius: 5px; display: flex; align-items: center; justify-content: center;">
          <span style="color: white; font-size: 12px; font-weight: bold;">M</span>
        </div>
        <span style="font-size: 14px; font-weight: 600; color: #111827;">Maintena</span>
      </div>
    </div>
    <div style="padding: 28px 24px;">
      <p style="font-size: 14px; color: #6b7280; margin: 0 0 16px 0;">${greeting}</p>
      <p style="font-size: 15px; color: #111827; margin: 0 0 8px 0; font-weight: 500;">A vendor has been assigned to your request.</p>
      <div style="background-color: #f9fafb; border-radius: 8px; padding: 16px; margin: 16px 0;">
        <p style="font-size: 13px; color: #374151; margin: 0 0 6px 0;"><strong>Request:</strong> ${escapeHtml(ticketTitle)}</p>
        <p style="font-size: 13px; color: #374151; margin: 0 0 6px 0;"><strong>Assigned vendor:</strong> ${escapeHtml(vendorName)}</p>
        ${propertyName ? `<p style="font-size: 13px; color: #374151; margin: 0;"><strong>Property:</strong> ${escapeHtml(propertyName)}</p>` : ''}
      </div>
      <p style="font-size: 13px; color: #6b7280; margin: 0;">Your property manager will keep you updated as work progresses.</p>
    </div>
    <div style="padding: 14px 24px; background-color: #f9fafb; border-top: 1px solid #f3f4f6;">
      <p style="font-size: 11px; color: #9ca3af; margin: 0;">© ${new Date().getFullYear()} Maintena · The AI operations layer for property maintenance.</p>
    </div>
  </div>
</body>
</html>`,
    })
    return !error
  } catch {
    return false
  }
}

export type SlaBreachItem = {
  id: string
  title: string
  urgency: string
  hoursOverdue: number
  propertyName?: string | null
  vendorName?: string | null
}

export async function sendSlaBreachEmail(
  managerEmail: string,
  items: SlaBreachItem[],
  appUrl: string
): Promise<boolean> {
  if (items.length === 0) return true
  const rows = items
    .map(
      it => `
      <tr>
        <td style="padding: 8px 0; border-bottom: 1px solid #f3f4f6;">
          <a href="${appUrl}/dashboard/tickets/${it.id}" style="font-size: 13px; color: #4f46e5; text-decoration: none; font-weight: 500;">${escapeHtml(it.title)}</a>
          <div style="font-size: 11px; color: #9ca3af; margin-top: 2px;">
            ${escapeHtml(it.urgency)} · ${it.hoursOverdue}h overdue${it.propertyName ? ` · ${escapeHtml(it.propertyName)}` : ''}${it.vendorName ? ` · ${escapeHtml(it.vendorName)}` : ' · unassigned'}
          </div>
        </td>
      </tr>`
    )
    .join('')
  try {
    const { error } = await resend.emails.send({
      from: 'Maintena <onboarding@resend.dev>',
      to: managerEmail,
      subject: `${items.length} maintenance ticket${items.length === 1 ? '' : 's'} past SLA`,
      html: `
<!DOCTYPE html>
<html>
<body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background-color: #f9fafb; padding: 40px 20px; margin: 0;">
  <div style="max-width: 520px; margin: 0 auto; background-color: #ffffff; border-radius: 12px; border: 1px solid #e5e7eb; overflow: hidden;">
    <div style="padding: 20px 24px; border-bottom: 1px solid #f3f4f6;">
      <span style="font-size: 14px; font-weight: 600; color: #111827;">Maintena</span>
    </div>
    <div style="padding: 24px;">
      <h2 style="font-size: 16px; font-weight: 600; color: #b45309; margin: 0 0 6px 0;">${items.length} ticket${items.length === 1 ? '' : 's'} need attention</h2>
      <p style="font-size: 13px; color: #6b7280; margin: 0 0 16px 0;">The following tickets have passed their SLA target and are not yet completed.</p>
      <table style="width: 100%; border-collapse: collapse;">${rows}</table>
    </div>
    <div style="padding: 14px 24px; background-color: #f9fafb; border-top: 1px solid #f3f4f6;">
      <p style="font-size: 11px; color: #9ca3af; margin: 0;">© ${new Date().getFullYear()} Maintena · You receive this because you manage this workspace.</p>
    </div>
  </div>
</body>
</html>`,
    })
    return !error
  } catch {
    return false
  }
}

type TenantStatusUpdateType = 'in_progress' | 'completed' | 'waiting'

const STATUS_COPY: Record<TenantStatusUpdateType, { subject: string; headline: string; body: string }> = {
  in_progress: {
    subject: 'Work has started on your maintenance request',
    headline: 'Work has started on your request.',
    body: 'A technician is now working on the repair. Your property manager will notify you when it is complete.',
  },
  completed: {
    subject: 'Your maintenance request has been completed',
    headline: 'Your repair has been completed.',
    body: 'The work has been finished and your ticket is now closed. Please contact your property manager if you have any concerns.',
  },
  waiting: {
    subject: 'Your maintenance request is on hold',
    headline: 'Your request is temporarily on hold.',
    body: 'The repair is currently waiting on parts or scheduling. Your property manager will follow up shortly.',
  },
}

export async function sendTenantStatusUpdateEmail(
  tenantEmail: string,
  tenantName: string | null,
  ticketTitle: string,
  newStatus: TenantStatusUpdateType,
  propertyName?: string | null
): Promise<boolean> {
  const copy = STATUS_COPY[newStatus]
  const greeting = tenantName ? `Hi ${escapeHtml(tenantName)},` : 'Hello,'
  try {
    const { error } = await resend.emails.send({
      from: 'Maintena <onboarding@resend.dev>',
      to: tenantEmail,
      subject: copy.subject,
      html: `
<!DOCTYPE html>
<html>
<body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background-color: #f9fafb; padding: 40px 20px; margin: 0;">
  <div style="max-width: 480px; margin: 0 auto; background-color: #ffffff; border-radius: 12px; border: 1px solid #e5e7eb; overflow: hidden;">
    <div style="padding: 20px 24px; border-bottom: 1px solid #f3f4f6;">
      <div style="display: inline-flex; align-items: center; gap: 8px;">
        <div style="width: 24px; height: 24px; background-color: #4f46e5; border-radius: 5px; display: flex; align-items: center; justify-content: center;">
          <span style="color: white; font-size: 12px; font-weight: bold;">M</span>
        </div>
        <span style="font-size: 14px; font-weight: 600; color: #111827;">Maintena</span>
      </div>
    </div>
    <div style="padding: 28px 24px;">
      <p style="font-size: 14px; color: #6b7280; margin: 0 0 16px 0;">${greeting}</p>
      <p style="font-size: 15px; color: #111827; margin: 0 0 8px 0; font-weight: 500;">${copy.headline}</p>
      <div style="background-color: #f9fafb; border-radius: 8px; padding: 16px; margin: 16px 0;">
        <p style="font-size: 13px; color: #374151; margin: 0 0 6px 0;"><strong>Request:</strong> ${escapeHtml(ticketTitle)}</p>
        ${propertyName ? `<p style="font-size: 13px; color: #374151; margin: 0;"><strong>Property:</strong> ${escapeHtml(propertyName)}</p>` : ''}
      </div>
      <p style="font-size: 13px; color: #6b7280; margin: 0;">${copy.body}</p>
    </div>
    <div style="padding: 14px 24px; background-color: #f9fafb; border-top: 1px solid #f3f4f6;">
      <p style="font-size: 11px; color: #9ca3af; margin: 0;">© ${new Date().getFullYear()} Maintena · The AI operations layer for property maintenance.</p>
    </div>
  </div>
</body>
</html>`,
    })
    return !error
  } catch {
    return false
  }
}

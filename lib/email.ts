import 'server-only'
import { Resend } from 'resend'

const resend = new Resend(process.env.RESEND_API_KEY)

export async function sendMagicLinkEmail(email: string, token: string): Promise<boolean> {
  const baseUrl = process.env.NEXT_PUBLIC_APP_URL || 'http://localhost:3000'
  const magicLink = `${baseUrl}/verify?token=${token}`

  try {
    const { error } = await resend.emails.send({
      from: 'Maintena <noreply@maintena.app>',
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
      from: 'Maintena <noreply@maintena.app>',
      to: vendorEmail,
      subject: `New job assigned: ${ticketTitle}`,
      html: `
<!DOCTYPE html>
<html>
<body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background-color: #f9fafb; padding: 40px 20px; margin: 0;">
  <div style="max-width: 480px; margin: 0 auto; background-color: #ffffff; border-radius: 12px; border: 1px solid #e5e7eb; padding: 32px 24px;">
    <p style="font-size: 15px; font-weight: 600; color: #111827; margin: 0 0 4px 0;">Maintena</p>
    <h2 style="font-size: 18px; font-weight: 600; color: #111827; margin: 16px 0 8px 0;">New job assigned to you</h2>
    <p style="font-size: 14px; color: #6b7280; margin: 0 0 16px 0;">Hi ${vendorName},</p>
    <p style="font-size: 14px; color: #374151; margin: 0 0 8px 0;"><strong>Job:</strong> ${ticketTitle}</p>
    ${propertyName ? `<p style="font-size: 14px; color: #374151; margin: 0 0 24px 0;"><strong>Property:</strong> ${propertyName}</p>` : ''}
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

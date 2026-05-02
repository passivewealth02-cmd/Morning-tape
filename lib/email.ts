import 'server-only'
import { Resend } from 'resend'

const resend = new Resend(process.env.RESEND_API_KEY)

export async function sendMagicLinkEmail(email: string, token: string): Promise<boolean> {
  const baseUrl = process.env.NEXT_PUBLIC_APP_URL || 'http://localhost:3000'
  const magicLink = `${baseUrl}/verify?token=${token}`

  try {
    const { error } = await resend.emails.send({
      from: 'The Morning Tape <noreply@themorningtape.com>',
      to: email,
      subject: 'Your Sign-In Link — The Morning Tape',
      html: `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
</head>
<body style="font-family: Georgia, serif; background-color: #F4EDE0; padding: 40px 20px; margin: 0;">
  <div style="max-width: 600px; margin: 0 auto; background-color: #F4EDE0;">
    <div style="text-align: center; padding-bottom: 24px; border-bottom: 1px solid #C9BFAB;">
      <h1 style="font-family: Georgia, serif; font-size: 28px; letter-spacing: 0.1em; text-transform: uppercase; color: #1A1612; margin: 0;">
        The Morning Tape
      </h1>
    </div>
    
    <div style="padding: 32px 0;">
      <p style="font-size: 18px; line-height: 1.6; color: #1A1612; margin: 0 0 24px 0;">
        You requested a sign-in link for your Morning Tape subscription.
      </p>
      
      <div style="text-align: center; margin: 32px 0;">
        <a href="${magicLink}" style="display: inline-block; background-color: #1A1612; color: #F4EDE0; padding: 16px 32px; text-decoration: none; font-family: Georgia, serif; letter-spacing: 0.05em;">
          Access Your Briefing
        </a>
      </div>
      
      <p style="font-size: 14px; color: #5C5549; line-height: 1.6; margin: 24px 0 0 0;">
        This link will expire in 15 minutes. If you did not request this email, you may safely ignore it.
      </p>
    </div>
    
    <div style="border-top: 1px solid #C9BFAB; padding-top: 24px; text-align: center;">
      <p style="font-size: 12px; color: #5C5549; margin: 0;">
        &copy; ${new Date().getFullYear()} The Morning Tape. All rights reserved.
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

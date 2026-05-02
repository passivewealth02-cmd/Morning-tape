# The Morning Tape

A premium AI-powered pre-market briefing delivered daily before the bell.

## Setup

### 1. Push to GitHub
Upload all files to your `morning-tape` GitHub repository.

### 2. Connect Vercel
- Go to vercel.com → "Add New Project"
- Import your `morning-tape` repo from GitHub
- Vercel will auto-detect Next.js

### 3. Set up Neon Database
- Go to console.neon.tech
- Create a new project
- Copy the connection string
- Open the SQL Editor and paste the contents of `init.sql` → Run
- Add `DATABASE_URL` to Vercel environment variables

### 4. Set up Stripe
- Go to dashboard.stripe.com → Products
- Create "The Trader" — $29/mo recurring
- Create "The Professional" — $49/mo recurring
- Copy both Price IDs
- Go to Developers → Webhooks → Add endpoint
- URL: `https://themorningtape.com/api/stripe/webhook`
- Events: `checkout.session.completed`, `customer.subscription.updated`, `customer.subscription.deleted`
- Copy the signing secret

### 5. Set up Resend
- Go to resend.com → API Keys → Create key
- Go to Domains → Add themorningtape.com → Add the DNS records at your domain registrar

### 6. Environment Variables
Add all variables from `.env.example` to Vercel → Settings → Environment Variables.

### 7. Connect Domain
In Vercel → Settings → Domains → Add `themorningtape.com` → Follow DNS instructions.

### 8. Deploy
Vercel auto-deploys when you push to GitHub. Redeploy after adding env variables.

## Pricing
- **The Trader** — $29/mo
- **The Professional** — $49/mo

export type LandingSection = {
  heading: string
  body: string[]
  bullets?: string[]
}

export type LandingPage = {
  slug: string
  metaTitle: string
  metaDescription: string
  keywords: string[]
  eyebrow: string
  h1: string
  intro: string
  ctaLabel: string
  sections: LandingSection[]
  benefits: { title: string; body: string }[]
  faqs: { q: string; a: string }[]
  related: string[]
}

export const LANDING_PAGES: LandingPage[] = [
  {
    slug: 'property-maintenance-software',
    metaTitle: 'Property Maintenance Software for Property Managers | Maintena',
    metaDescription:
      'Property maintenance software that captures requests, triages them with AI, dispatches vendors, and tracks every repair to completion. Start free with Maintena.',
    keywords: [
      'property maintenance software',
      'maintenance management software',
      'property management maintenance platform',
      'landlord maintenance software',
      'property operations software',
    ],
    eyebrow: 'Property maintenance software',
    h1: 'Property maintenance software built for modern property managers',
    intro:
      'Maintena is property maintenance software that turns scattered tenant requests, phone calls, and emails into one organized workflow. Capture every maintenance request, let AI triage and prioritize it, dispatch the right vendor, and track the repair to completion — without spreadsheets or sticky notes.',
    ctaLabel: 'Start managing maintenance free',
    sections: [
      {
        heading: 'Why property managers switch to dedicated maintenance software',
        body: [
          'Most property teams still coordinate repairs across email threads, text messages, and a spreadsheet that only one person understands. Requests fall through the cracks, vendors are double-booked, and tenants are left wondering whether anyone heard them. As a portfolio grows, that manual coordination becomes the single biggest source of wasted hours and frustrated residents.',
          'Maintenance management software centralizes the entire repair lifecycle in one place. Every request becomes a structured ticket with an owner, a status, a property, and a clear history. Instead of reconstructing what happened from memory, your team works from a shared source of truth that updates in real time.',
        ],
      },
      {
        heading: 'Everything you need to run property maintenance in one platform',
        body: [
          'Maintena combines intake, triage, dispatch, tracking, and reporting so you never have to stitch together half a dozen tools. It is property operations software designed around how maintenance actually flows — from the moment a tenant reports an issue to the final invoice.',
        ],
        bullets: [
          'Automatic request capture from email, web form, and SMS — no tenant app required',
          'AI categorization that detects trade, urgency, and safety risk in seconds',
          'Smart vendor dispatch ranked by trade, location, availability, and past performance',
          'A live Kanban board and sortable table for every open ticket',
          'SLA timers that warn you before a ticket goes overdue',
          'A timestamped audit trail for full accountability',
        ],
      },
      {
        heading: 'Built to scale from a single building to a full portfolio',
        body: [
          'Whether you manage one property or hundreds, Maintena keeps coordination consistent. Filter tickets by property, urgency, or vendor, and give every stakeholder — owners, on-site staff, and vendors — exactly the visibility they need.',
          'Because the platform is fast and runs in the browser, your team can work from the office, the field, or their phone without installing anything. New team members are productive on day one because the workflow is the same for every request.',
        ],
      },
    ],
    benefits: [
      { title: 'Stop losing requests', body: 'Every email, form submission, and text becomes a tracked ticket so nothing slips through.' },
      { title: 'Cut response times', body: 'AI triage and one-click dispatch get the right vendor moving in minutes, not days.' },
      { title: 'Keep tenants informed', body: 'Automated updates at each stage reduce inbound "any update?" calls.' },
      { title: 'Prove accountability', body: 'A full audit timeline shows who did what and when, for every repair.' },
    ],
    faqs: [
      {
        q: 'What is property maintenance software?',
        a: 'Property maintenance software is a platform that helps property managers capture maintenance requests, assign and dispatch vendors, track repair progress, and report on performance — replacing the spreadsheets, email threads, and phone calls that teams typically rely on.',
      },
      {
        q: 'Who is Maintena for?',
        a: 'Maintena is built for residential and commercial property managers, landlords, and property operations teams who coordinate repairs across one or many properties and want to reduce manual work while improving response times.',
      },
      {
        q: 'Do tenants need to download an app?',
        a: 'No. Tenants can submit requests by email, web form, or SMS. Maintena captures the request automatically and turns it into a structured ticket, so there is nothing for residents to install or learn.',
      },
      {
        q: 'How long does it take to get started?',
        a: 'Most teams are set up in under five minutes. You can start free, add your properties and vendors, and begin routing requests the same day — no credit card required.',
      },
    ],
    related: ['vendor-dispatch-software', 'maintenance-ticket-management', 'tenant-maintenance-requests'],
  },
  {
    slug: 'vendor-dispatch-software',
    metaTitle: 'Vendor Dispatch Software for Maintenance Teams | Maintena',
    metaDescription:
      'Vendor dispatch software that recommends the best contractor by trade, location, and performance — then assigns and notifies them in one click. Try Maintena free.',
    keywords: [
      'vendor dispatch software',
      'vendor management platform',
      'maintenance coordination platform',
      'repair coordination software',
    ],
    eyebrow: 'Vendor dispatch software',
    h1: 'Vendor dispatch software that picks the right contractor automatically',
    intro:
      'Maintena is vendor dispatch software that ranks your contractors by trade, location, availability, and track record — then lets you assign and notify the best match in a single click. No more cold-calling vendors or wondering who is available.',
    ctaLabel: 'Dispatch your first vendor free',
    sections: [
      {
        heading: 'Turn vendor coordination from a phone tree into one click',
        body: [
          'Dispatching the right vendor usually means scrolling through a contact list, calling around to check availability, and re-explaining the same job to each one. It is slow, it is error-prone, and it gets harder every time you add a property or a trade.',
          'Maintena replaces that scramble with an intelligent recommendation. For each ticket, it surfaces the vendors most likely to do a great job quickly, complete with a match score you can understand. You stay in control of the final decision while the software does the legwork.',
        ],
      },
      {
        heading: 'How AI vendor matching works',
        body: [
          'Every time a request comes in, Maintena reads the issue, determines the required trade, and scores your vendor network against the job. The ranking balances the factors that actually predict a good outcome.',
        ],
        bullets: [
          'Trade fit — plumbers for leaks, electricians for wiring, HVAC techs for heating',
          'Proximity to the property to reduce travel time and trip fees',
          'Current availability so you do not assign someone already booked solid',
          'Historical performance, ratings, and on-time completion rate',
        ],
      },
      {
        heading: 'Notify vendors with everything they need — automatically',
        body: [
          'Once you assign a vendor, Maintena sends them the full job context: the property address, the tenant report, photos, urgency, and any access notes. There is no back-and-forth to gather details, so work starts faster.',
          'As a maintenance coordination platform, Maintena also keeps the tenant and your team in the loop, so everyone knows the job is assigned and when to expect the vendor — without anyone making a single status call.',
        ],
      },
    ],
    benefits: [
      { title: 'Faster dispatch', body: 'Assign and notify the best-matched vendor in one click instead of a round of phone calls.' },
      { title: 'Better outcomes', body: 'Matching by performance keeps your most reliable vendors on the most important jobs.' },
      { title: 'Fewer mistakes', body: 'Vendors get complete job details automatically, so nothing is missed on the first visit.' },
      { title: 'Balanced workload', body: 'Distribute work intelligently across your network instead of overloading one or two contractors.' },
    ],
    faqs: [
      {
        q: 'What is vendor dispatch software?',
        a: 'Vendor dispatch software helps maintenance teams choose, assign, and notify the right contractor for each job. Maintena adds AI matching that ranks vendors by trade, location, availability, and past performance so dispatch takes seconds.',
      },
      {
        q: 'Can I keep my existing vendors?',
        a: 'Yes. Add your current vendor network to Maintena and the platform will rank them for each job. You are never locked into a marketplace — your vendors stay your vendors.',
      },
      {
        q: 'Does the vendor need an account?',
        a: 'Vendors receive job details and updates automatically by email, so they can get started without setting up anything. You stay in control of assignments and communication.',
      },
      {
        q: 'Can I override the AI recommendation?',
        a: 'Always. Maintena recommends the best match, but you make the final call and can assign any vendor you choose.',
      },
    ],
    related: ['property-maintenance-software', 'maintenance-ticket-management', 'ai-property-management-tools'],
  },
  {
    slug: 'maintenance-ticket-management',
    metaTitle: 'Maintenance Ticket Tracking & Management Software | Maintena',
    metaDescription:
      'Track every maintenance ticket on a live Kanban board with SLA timers, status updates, and a full audit trail. Maintenance ticket management made simple with Maintena.',
    keywords: [
      'maintenance ticket tracking',
      'maintenance ticket management',
      'maintenance workflow automation',
      'property maintenance workflow',
    ],
    eyebrow: 'Maintenance ticket management',
    h1: 'Maintenance ticket tracking that keeps every repair on schedule',
    intro:
      'Maintena gives every maintenance request a structured ticket and a clear status, then tracks it on a live board until it is resolved. SLA timers, automated updates, and a complete audit trail mean nothing stalls and nothing is forgotten.',
    ctaLabel: 'Track your tickets free',
    sections: [
      {
        heading: 'See every open ticket at a glance',
        body: [
          'When repairs live in someone’s inbox, it is impossible to know what is actually happening across your portfolio. Maintena puts every ticket on a real-time Kanban board — New, Assigned, In Progress, and Completed — so you can read the state of your operation in seconds.',
          'Prefer a list? Switch to a sortable table and filter by property, urgency, or vendor. The data is the same; you just choose the view that fits the task in front of you.',
        ],
      },
      {
        heading: 'Maintenance workflow automation that prevents overdue tickets',
        body: [
          'Set response and resolution targets, and Maintena watches the clock for you. SLA timers flag tickets that are approaching their deadline so your team can intervene before a tenant has to follow up.',
        ],
        bullets: [
          'Automatic status transitions as work progresses',
          'SLA alerts before a ticket goes overdue',
          'Automated tenant and vendor updates at each stage',
          'Filters and saved views for properties, trades, and priorities',
        ],
      },
      {
        heading: 'A complete, timestamped audit trail',
        body: [
          'Every action on a ticket — created, assigned, updated, completed — is logged with a timestamp and an owner. When an owner asks what happened on a repair, or you need a record for compliance, the full history is right there.',
          'That accountability is what turns reactive firefighting into a measurable process. Over time, ticket data shows you where bottlenecks are, which properties generate the most work, and which vendors deliver.',
        ],
      },
    ],
    benefits: [
      { title: 'Total visibility', body: 'A live board shows the status of every repair across your whole portfolio.' },
      { title: 'Nothing overdue', body: 'SLA timers warn you before deadlines slip so tickets never stall silently.' },
      { title: 'Less follow-up', body: 'Automated updates keep tenants and vendors informed without manual messages.' },
      { title: 'Clear records', body: 'A timestamped audit trail documents every action for owners and compliance.' },
    ],
    faqs: [
      {
        q: 'What is maintenance ticket management?',
        a: 'Maintenance ticket management is the process of capturing each repair request as a trackable ticket and moving it through clear stages — intake, assignment, work, and completion — with status, ownership, and history recorded along the way.',
      },
      {
        q: 'How does SLA tracking work in Maintena?',
        a: 'You set response and resolution targets, and Maintena starts a timer on each ticket. As a ticket nears its target it is flagged so your team can act before it becomes overdue.',
      },
      {
        q: 'Can I see tickets by property or vendor?',
        a: 'Yes. Filter and sort tickets by property, urgency, vendor, or status in both the Kanban and table views to focus on exactly what you need.',
      },
      {
        q: 'Is there a record of changes?',
        a: 'Every ticket has a complete, timestamped audit trail showing who created, assigned, updated, and completed it — useful for accountability and compliance.',
      },
    ],
    related: ['property-maintenance-software', 'vendor-dispatch-software', 'tenant-maintenance-requests'],
  },
  {
    slug: 'tenant-maintenance-requests',
    metaTitle: 'Tenant Maintenance Request Software & Automation | Maintena',
    metaDescription:
      'Let tenants report issues by email, web form, or SMS — no app required. Maintena turns every tenant maintenance request into a tracked ticket automatically. Try it free.',
    keywords: [
      'tenant maintenance requests',
      'maintenance request automation',
      'maintenance request management',
      'tenant repair requests',
    ],
    eyebrow: 'Tenant maintenance requests',
    h1: 'Tenant maintenance requests, captured and resolved automatically',
    intro:
      'Maintena makes it effortless for tenants to report a problem and effortless for you to act on it. Residents reach out by email, web form, or SMS — no app to download — and every request becomes a structured, tracked ticket the moment it arrives.',
    ctaLabel: 'Streamline tenant requests free',
    sections: [
      {
        heading: 'Meet tenants where they already are',
        body: [
          'Asking residents to learn a portal or download an app is the fastest way to get requests by text at midnight instead. Maintena removes that friction entirely. Tenants report issues the way they already communicate, and the platform does the structuring on your side.',
          'A forwarded email address, a simple web form, or an SMS is all it takes. Photos and details come through attached to the ticket, so your team has the context it needs from the first glance.',
        ],
      },
      {
        heading: 'Maintenance request automation from intake to update',
        body: [
          'Every incoming request is automatically captured, categorized by AI, and prioritized by urgency. Emergencies — a burst pipe, a sparking outlet, no heat in winter — are flagged immediately so they never wait in a queue behind a light-bulb change.',
        ],
        bullets: [
          'Instant capture from email, web form, and SMS',
          'AI categorization by trade and urgency',
          'Automatic acknowledgement so tenants know they were heard',
          'Status updates at each stage through to completion',
        ],
      },
      {
        heading: 'Happier residents, fewer status calls',
        body: [
          'A huge share of property management calls are simply "what is happening with my repair?" Maintena answers that question before it is asked. Tenants get a confirmation when the request is received, a heads-up when a vendor is on the way, and a note when the job is done.',
          'That steady, professional communication builds trust, improves renewals, and frees your team from the endless follow-up that comes with manual coordination.',
        ],
      },
    ],
    benefits: [
      { title: 'No tenant app', body: 'Residents report issues by email, form, or SMS — nothing to download or learn.' },
      { title: 'Instant capture', body: 'Every request becomes a structured ticket automatically, with photos attached.' },
      { title: 'Emergencies first', body: 'AI flags urgent and safety issues so they jump the queue.' },
      { title: 'Automatic updates', body: 'Tenants stay informed at every stage, cutting "any update?" calls.' },
    ],
    faqs: [
      {
        q: 'How do tenants submit maintenance requests?',
        a: 'Tenants can submit requests by email, a simple web form, or SMS. Maintena captures each one automatically and converts it into a structured ticket — no app or portal login required.',
      },
      {
        q: 'Are urgent requests prioritized?',
        a: 'Yes. AI reads each request and flags emergencies — like leaks, electrical hazards, or no heat — as high urgency so they are addressed before routine work.',
      },
      {
        q: 'Do tenants get updates on their request?',
        a: 'Tenants receive automatic updates at each stage: confirmation that the request was received, notice when a vendor is assigned and en route, and a note when the job is complete.',
      },
      {
        q: 'Can tenants attach photos?',
        a: 'Yes. Photos and details submitted with a request are attached to the ticket so your team and vendors have full context from the start.',
      },
    ],
    related: ['property-maintenance-software', 'maintenance-ticket-management', 'ai-property-management-tools'],
  },
  {
    slug: 'ai-property-management-tools',
    metaTitle: 'AI Property Management & Maintenance Tools | Maintena',
    metaDescription:
      'AI property maintenance software that classifies requests, recommends vendors, and automates updates. See how AI maintenance automation saves hours every week with Maintena.',
    keywords: [
      'ai property maintenance software',
      'ai property management tools',
      'ai maintenance automation',
      'maintenance workflow automation',
    ],
    eyebrow: 'AI property management tools',
    h1: 'AI property maintenance software that does the busywork for you',
    intro:
      'Maintena applies AI to the most time-consuming parts of property maintenance — reading requests, setting priority, matching vendors, and keeping everyone updated — so your team can focus on decisions instead of data entry.',
    ctaLabel: 'See AI maintenance in action',
    sections: [
      {
        heading: 'Where AI actually helps in property maintenance',
        body: [
          'AI is only useful when it removes real work. In property maintenance, the repetitive, judgment-light tasks add up fast: reading every incoming message, deciding the trade, gauging urgency, finding an available vendor, and sending the same updates over and over. Maintena automates exactly those steps.',
          'The result is not a gimmick — it is hours back every week and faster, more consistent service for tenants. Your team still makes the calls that matter; the AI handles the repetitive groundwork that used to eat the day.',
        ],
      },
      {
        heading: 'AI maintenance automation, step by step',
        body: [
          'From intake to resolution, AI works behind the scenes at each stage so coordination happens almost on its own.',
        ],
        bullets: [
          'Classification — reads each request and assigns the right trade and category',
          'Prioritization — scores urgency and flags safety risks automatically',
          'Vendor matching — ranks contractors by fit, location, availability, and performance',
          'Summaries — writes a plain-English summary of each issue for your team',
          'Updates — keeps tenants and vendors informed at every milestone',
        ],
      },
      {
        heading: 'Consistent, 24/7 coordination',
        body: [
          'People get tired, take vacations, and prioritize differently. AI applies the same logic to every request, at any hour, so a 2 a.m. emergency is triaged with the same rigor as a 2 p.m. routine job.',
          'That consistency is what makes maintenance workflow automation dependable. As your portfolio grows, the system scales without adding coordination headcount — and every decision is logged for full transparency.',
        ],
      },
    ],
    benefits: [
      { title: 'Hours saved weekly', body: 'AI handles reading, sorting, and routing so your team skips the busywork.' },
      { title: 'Consistent triage', body: 'Every request is prioritized with the same logic, around the clock.' },
      { title: 'Smarter dispatch', body: 'Vendor recommendations improve outcomes by matching the right pro to each job.' },
      { title: 'Scales with you', body: 'Automation absorbs more volume without more coordination headcount.' },
    ],
    faqs: [
      {
        q: 'What is AI property maintenance software?',
        a: 'AI property maintenance software uses artificial intelligence to automate maintenance coordination — reading and classifying requests, scoring urgency, recommending vendors, and keeping everyone updated — so property teams spend less time on manual work.',
      },
      {
        q: 'Will AI replace my team?',
        a: 'No. Maintena automates repetitive groundwork like triage and routing so your team can focus on decisions, relationships, and quality. People stay in control of every important call.',
      },
      {
        q: 'How accurate is the AI classification?',
        a: 'Maintena’s AI reliably identifies trade and urgency for typical maintenance requests and flags safety issues, and you can always review or adjust its recommendations before acting.',
      },
      {
        q: 'Is my data secure?',
        a: 'Maintena keeps a complete, timestamped audit trail of every action and is built so your team controls access and assignments. Your vendor network and tenant data stay yours.',
      },
    ],
    related: ['property-maintenance-software', 'vendor-dispatch-software', 'maintenance-ticket-management'],
  },
]

export function getLandingPage(slug: string) {
  return LANDING_PAGES.find(p => p.slug === slug)
}

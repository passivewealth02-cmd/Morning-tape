-- Maintena: AI Maintenance & Vendor Coordination Platform
-- Run this SQL in your Neon database console

-- Organizations (multi-tenant root)
CREATE TABLE IF NOT EXISTS organizations (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  name TEXT NOT NULL,
  slug TEXT UNIQUE NOT NULL,
  inbox_token TEXT UNIQUE,
  plan TEXT NOT NULL DEFAULT 'trial' CHECK (plan IN ('trial', 'starter', 'growth', 'pro')),
  plan_status TEXT NOT NULL DEFAULT 'active' CHECK (plan_status IN ('active', 'past_due', 'canceled', 'trialing', 'incomplete')),
  trial_ends_at TIMESTAMP,
  stripe_customer_id TEXT,
  stripe_subscription_id TEXT,
  stripe_price_id TEXT,
  created_at TIMESTAMP DEFAULT NOW(),
  updated_at TIMESTAMP DEFAULT NOW()
);

-- Users
CREATE TABLE IF NOT EXISTS users (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  organization_id UUID REFERENCES organizations(id) ON DELETE CASCADE,
  email TEXT UNIQUE NOT NULL,
  name TEXT,
  role TEXT NOT NULL DEFAULT 'manager' CHECK (role IN ('admin', 'manager', 'coordinator', 'vendor')),
  created_at TIMESTAMP DEFAULT NOW(),
  updated_at TIMESTAMP DEFAULT NOW()
);

-- Magic link tokens
CREATE TABLE IF NOT EXISTS magic_link_tokens (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  email TEXT NOT NULL,
  token TEXT UNIQUE NOT NULL,
  expires_at TIMESTAMP NOT NULL,
  used BOOLEAN DEFAULT FALSE,
  created_at TIMESTAMP DEFAULT NOW()
);

-- Sessions
CREATE TABLE IF NOT EXISTS sessions (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  user_id UUID NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  token TEXT UNIQUE NOT NULL,
  expires_at TIMESTAMP NOT NULL,
  created_at TIMESTAMP DEFAULT NOW()
);

-- Properties
CREATE TABLE IF NOT EXISTS properties (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  organization_id UUID NOT NULL REFERENCES organizations(id) ON DELETE CASCADE,
  name TEXT NOT NULL,
  address TEXT NOT NULL,
  city TEXT,
  province TEXT,
  unit_count INTEGER DEFAULT 0,
  created_at TIMESTAMP DEFAULT NOW(),
  updated_at TIMESTAMP DEFAULT NOW()
);

-- Units
CREATE TABLE IF NOT EXISTS units (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  property_id UUID NOT NULL REFERENCES properties(id) ON DELETE CASCADE,
  unit_number TEXT NOT NULL,
  tenant_name TEXT,
  tenant_email TEXT,
  tenant_phone TEXT,
  created_at TIMESTAMP DEFAULT NOW()
);

-- Vendors
CREATE TABLE IF NOT EXISTS vendors (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  organization_id UUID NOT NULL REFERENCES organizations(id) ON DELETE CASCADE,
  name TEXT NOT NULL,
  trade_type TEXT NOT NULL,
  email TEXT,
  phone TEXT,
  service_areas TEXT[],
  availability TEXT DEFAULT 'available' CHECK (availability IN ('available', 'busy', 'unavailable')),
  rating NUMERIC(2,1) DEFAULT 0,
  insurance_status TEXT DEFAULT 'unknown' CHECK (insurance_status IN ('verified', 'expired', 'unknown')),
  notes TEXT,
  created_at TIMESTAMP DEFAULT NOW(),
  updated_at TIMESTAMP DEFAULT NOW()
);

-- Maintenance tickets
CREATE TABLE IF NOT EXISTS maintenance_tickets (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  organization_id UUID NOT NULL REFERENCES organizations(id) ON DELETE CASCADE,
  property_id UUID REFERENCES properties(id),
  unit_id UUID REFERENCES units(id),
  title TEXT NOT NULL,
  description TEXT NOT NULL,
  status TEXT NOT NULL DEFAULT 'new' CHECK (status IN ('new', 'assigned', 'in_progress', 'waiting', 'completed', 'cancelled')),
  urgency TEXT NOT NULL DEFAULT 'medium' CHECK (urgency IN ('low', 'medium', 'high', 'emergency')),
  ai_category TEXT,
  ai_vendor_type TEXT,
  ai_summary TEXT,
  ai_escalation_risk BOOLEAN DEFAULT FALSE,
  assigned_vendor_id UUID REFERENCES vendors(id),
  created_by UUID REFERENCES users(id),
  tenant_name TEXT,
  tenant_email TEXT,
  tenant_phone TEXT,
  first_response_at TIMESTAMP,
  assigned_at TIMESTAMP,
  completed_at TIMESTAMP,
  sla_due_at TIMESTAMP,
  created_at TIMESTAMP DEFAULT NOW(),
  updated_at TIMESTAMP DEFAULT NOW()
);

-- Ticket messages / notes
CREATE TABLE IF NOT EXISTS ticket_messages (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  ticket_id UUID NOT NULL REFERENCES maintenance_tickets(id) ON DELETE CASCADE,
  sender_type TEXT NOT NULL CHECK (sender_type IN ('manager', 'vendor', 'tenant', 'system')),
  sender_id UUID REFERENCES users(id),
  message TEXT NOT NULL,
  is_internal BOOLEAN DEFAULT FALSE,
  created_at TIMESTAMP DEFAULT NOW()
);

-- Ticket files / photos
CREATE TABLE IF NOT EXISTS ticket_files (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  ticket_id UUID NOT NULL REFERENCES maintenance_tickets(id) ON DELETE CASCADE,
  file_url TEXT NOT NULL,
  file_name TEXT,
  file_type TEXT,
  uploaded_by UUID REFERENCES users(id),
  created_at TIMESTAMP DEFAULT NOW()
);

-- Activity logs (audit timeline)
CREATE TABLE IF NOT EXISTS activity_logs (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  organization_id UUID NOT NULL REFERENCES organizations(id) ON DELETE CASCADE,
  ticket_id UUID REFERENCES maintenance_tickets(id),
  user_id UUID REFERENCES users(id),
  action_type TEXT NOT NULL,
  description TEXT NOT NULL,
  metadata JSONB,
  created_at TIMESTAMP DEFAULT NOW()
);

-- Performance indexes
CREATE INDEX IF NOT EXISTS idx_users_organization_id ON users(organization_id);
CREATE INDEX IF NOT EXISTS idx_sessions_token ON sessions(token);
CREATE INDEX IF NOT EXISTS idx_magic_link_tokens_token ON magic_link_tokens(token);
CREATE INDEX IF NOT EXISTS idx_properties_organization_id ON properties(organization_id);
CREATE INDEX IF NOT EXISTS idx_vendors_organization_id ON vendors(organization_id);
CREATE INDEX IF NOT EXISTS idx_tickets_organization_id ON maintenance_tickets(organization_id);
CREATE INDEX IF NOT EXISTS idx_tickets_status ON maintenance_tickets(status);
CREATE INDEX IF NOT EXISTS idx_tickets_urgency ON maintenance_tickets(urgency);
CREATE INDEX IF NOT EXISTS idx_tickets_assigned_vendor ON maintenance_tickets(assigned_vendor_id);
CREATE INDEX IF NOT EXISTS idx_ticket_messages_ticket_id ON ticket_messages(ticket_id);
CREATE INDEX IF NOT EXISTS idx_activity_logs_ticket_id ON activity_logs(ticket_id);
CREATE INDEX IF NOT EXISTS idx_activity_logs_organization_id ON activity_logs(organization_id);

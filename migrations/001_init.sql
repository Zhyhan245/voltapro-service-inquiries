-- Початкова схема бази даних корпоративного сайту «ВольтаПро Сервіс»

CREATE TABLE IF NOT EXISTS schema_migrations (
  version TEXT PRIMARY KEY,
  applied_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS services (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  slug TEXT NOT NULL UNIQUE,
  title TEXT NOT NULL,
  category TEXT NOT NULL CHECK (category IN ('Проєктування', 'Монтаж', 'Обслуговування', 'Аудит')),
  short_description TEXT NOT NULL,
  full_description TEXT NOT NULL,
  base_price_uah INTEGER NOT NULL CHECK (base_price_uah >= 0),
  sla_business_days INTEGER NOT NULL CHECK (sla_business_days BETWEEN 1 AND 30),
  is_active INTEGER NOT NULL DEFAULT 1 CHECK (is_active IN (0, 1)),
  created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS employees (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  full_name TEXT NOT NULL,
  email TEXT NOT NULL UNIQUE,
  password_hash TEXT NOT NULL,
  role TEXT NOT NULL CHECK (role IN ('MANAGER', 'SENIOR_MANAGER')),
  department TEXT NOT NULL,
  is_active INTEGER NOT NULL DEFAULT 1 CHECK (is_active IN (0, 1)),
  created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS client_contacts (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  contact_name TEXT NOT NULL,
  company_name TEXT,
  phone TEXT NOT NULL,
  email TEXT NOT NULL,
  created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS inquiries (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  ticket_number TEXT NOT NULL UNIQUE,
  access_code TEXT NOT NULL,
  service_id INTEGER NOT NULL REFERENCES services(id),
  contact_id INTEGER NOT NULL REFERENCES client_contacts(id),
  assigned_employee_id INTEGER REFERENCES employees(id),
  subject TEXT NOT NULL,
  message TEXT NOT NULL,
  status TEXT NOT NULL CHECK (status IN ('NEW', 'IN_PROGRESS', 'COMPLETED', 'REJECTED')),
  rejection_reason TEXT,
  public_response TEXT,
  idempotency_key TEXT UNIQUE,
  sla_business_days INTEGER NOT NULL CHECK (sla_business_days BETWEEN 1 AND 30),
  due_date TEXT NOT NULL,
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL,
  closed_at TEXT
);

CREATE TABLE IF NOT EXISTS inquiry_history (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  inquiry_id INTEGER NOT NULL REFERENCES inquiries(id) ON DELETE CASCADE,
  author_employee_id INTEGER REFERENCES employees(id),
  event_type TEXT NOT NULL CHECK (event_type IN ('CREATED', 'STATUS_CHANGE', 'ASSIGNMENT', 'PUBLIC_COMMENT', 'INTERNAL_NOTE')),
  old_status TEXT,
  new_status TEXT,
  comment TEXT NOT NULL,
  is_internal INTEGER NOT NULL DEFAULT 0 CHECK (is_internal IN (0, 1)),
  created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS calendar_exceptions (
  date TEXT PRIMARY KEY,
  is_working_day INTEGER NOT NULL CHECK (is_working_day IN (0, 1)),
  description TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS sessions (
  token TEXT PRIMARY KEY,
  employee_id INTEGER NOT NULL REFERENCES employees(id) ON DELETE CASCADE,
  expires_at TEXT NOT NULL,
  created_at TEXT NOT NULL
);

-- Індекси для швидкого пошуку, фільтрації та звіту прострочень SLA
CREATE INDEX IF NOT EXISTS idx_inquiries_status_due ON inquiries(status, due_date);
CREATE INDEX IF NOT EXISTS idx_inquiries_assigned_status ON inquiries(assigned_employee_id, status);
CREATE INDEX IF NOT EXISTS idx_inquiries_created_at ON inquiries(created_at DESC);
CREATE INDEX IF NOT EXISTS idx_inquiry_history_inquiry_id ON inquiry_history(inquiry_id, created_at);
CREATE INDEX IF NOT EXISTS idx_sessions_expires_at ON sessions(expires_at);

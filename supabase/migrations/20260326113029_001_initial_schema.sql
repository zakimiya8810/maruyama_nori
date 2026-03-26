/*
  # Initial Schema for Customer Karte Application

  1. New Tables
    - `departments` - 部門マスタ
      - `id` (uuid, primary key)
      - `code` (text, unique) - 部門コード
      - `name` (text) - 部門名
      - `division` (text) - 大区分
      - `display_order` (integer) - 表示順
    
    - `employees` - 社員マスタ
      - `id` (uuid, primary key)
      - `employee_code` (text, unique) - 担当者コード
      - `login_id` (text, unique) - ログインID
      - `password_hash` (text) - パスワードハッシュ
      - `name` (text) - 担当者名
      - `email` (text) - メールアドレス
      - `role` (text) - 役職
      - `department_id` (uuid) - 部門ID
      - `admin_permission` (text) - 管理権限
      - `is_retired` (boolean) - 退職者フラグ
      - `display_order` (integer) - 表示順
    
    - `ranks` - ランクマスタ
      - `id` (uuid, primary key)
      - `code` (text, unique) - 名称コード
      - `name` (text) - ランク名称
      - `visibility` (text) - 表示有無
      - `display_order` (integer) - 表示順
    
    - `business_types` - 業態マスタ
      - `id` (uuid, primary key)
      - `code` (text, unique) - 業態CD
      - `name` (text) - 業態名
    
    - `payment_terms` - 締日マスタ
      - `id` (uuid, primary key)
      - `pattern` (text) - 締日パターン
      - `display_order` (integer) - 表示順
    
    - `customers` - 得意先マスタ
      - Full customer info including billing, address, etc.
    
    - `products` - 商品マスタ
      - `id` (uuid, primary key)
      - `product_code` (text, unique) - 商品コード
      - `name` (text) - 商品名
      - `category` (text) - カテゴリ
      - `standard_price` (numeric) - 標準単価
    
    - `customer_prices` - 単価マスタ
      - Customer-specific pricing
    
    - `meetings` - 商談管理
      - Meeting/negotiation records
    
    - `applications` - 申請管理
      - Approval workflow applications
    
    - `application_customer_data` - 申請データ_顧客
    - `application_price_data` - 申請データ_単価
    
    - `jutaikin` - 渋滞金リスト
    - `customer_sales` - 売上データ

  2. Security
    - Enable RLS on all tables
    - Policies for authenticated users
*/

-- ====================================
-- 部門マスタ
-- ====================================
CREATE TABLE IF NOT EXISTS departments (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  code text UNIQUE NOT NULL,
  name text NOT NULL,
  division text DEFAULT '',
  display_order integer DEFAULT 999,
  created_at timestamptz DEFAULT now(),
  updated_at timestamptz DEFAULT now()
);

ALTER TABLE departments ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Authenticated users can view departments"
  ON departments FOR SELECT
  TO authenticated
  USING (true);

-- ====================================
-- 社員マスタ
-- ====================================
CREATE TABLE IF NOT EXISTS employees (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  employee_code text UNIQUE,
  login_id text UNIQUE NOT NULL,
  password_hash text NOT NULL,
  name text NOT NULL,
  email text DEFAULT '',
  role text DEFAULT '',
  department_id uuid REFERENCES departments(id),
  division text DEFAULT '',
  admin_permission text DEFAULT '',
  is_retired boolean DEFAULT false,
  display_order integer DEFAULT 999,
  created_at timestamptz DEFAULT now(),
  updated_at timestamptz DEFAULT now()
);

ALTER TABLE employees ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Authenticated users can view employees"
  ON employees FOR SELECT
  TO authenticated
  USING (true);

CREATE POLICY "Users can update own profile"
  ON employees FOR UPDATE
  TO authenticated
  USING (id = auth.uid())
  WITH CHECK (id = auth.uid());

-- ====================================
-- ランクマスタ
-- ====================================
CREATE TABLE IF NOT EXISTS ranks (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  code text UNIQUE,
  name text NOT NULL,
  visibility text DEFAULT '表示',
  display_order integer DEFAULT 999,
  created_at timestamptz DEFAULT now()
);

ALTER TABLE ranks ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Authenticated users can view ranks"
  ON ranks FOR SELECT
  TO authenticated
  USING (true);

-- ====================================
-- 業態マスタ
-- ====================================
CREATE TABLE IF NOT EXISTS business_types (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  code text UNIQUE NOT NULL,
  name text NOT NULL,
  created_at timestamptz DEFAULT now()
);

ALTER TABLE business_types ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Authenticated users can view business types"
  ON business_types FOR SELECT
  TO authenticated
  USING (true);

-- ====================================
-- 締日マスタ
-- ====================================
CREATE TABLE IF NOT EXISTS payment_terms (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  pattern text NOT NULL,
  display_order integer DEFAULT 999,
  created_at timestamptz DEFAULT now()
);

ALTER TABLE payment_terms ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Authenticated users can view payment terms"
  ON payment_terms FOR SELECT
  TO authenticated
  USING (true);

-- ====================================
-- 得意先マスタ
-- ====================================
CREATE TABLE IF NOT EXISTS customers (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  customer_code text UNIQUE NOT NULL,
  name text NOT NULL,
  short_name text DEFAULT '',
  kana_name text DEFAULT '',
  postal_code text DEFAULT '',
  address_1 text DEFAULT '',
  address_2 text DEFAULT '',
  tel text DEFAULT '',
  fax text DEFAULT '',
  business_type_id uuid REFERENCES business_types(id),
  rank_id uuid REFERENCES ranks(id),
  handler_employee_id uuid REFERENCES employees(id),
  billing_customer_id uuid,
  payment_term_id uuid REFERENCES payment_terms(id),
  tax_method text DEFAULT '',
  credit_limit numeric DEFAULT 0,
  shipping_location text DEFAULT '',
  delivery_method text DEFAULT '',
  group_code text DEFAULT '',
  representative_name text DEFAULT '',
  contact_person text DEFAULT '',
  nori_usage numeric DEFAULT 0,
  tea_usage numeric DEFAULT 0,
  natural_food_handling text DEFAULT '',
  shipping_fee text DEFAULT '',
  registration_date date,
  trade_start_date date,
  delete_flag integer DEFAULT 0,
  billing_postal_code text DEFAULT '',
  billing_name text DEFAULT '',
  billing_address_1 text DEFAULT '',
  billing_address_2 text DEFAULT '',
  billing_tel text DEFAULT '',
  billing_fax text DEFAULT '',
  last_updated timestamptz DEFAULT now(),
  created_at timestamptz DEFAULT now(),
  updated_at timestamptz DEFAULT now()
);

ALTER TABLE customers ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Authenticated users can view customers"
  ON customers FOR SELECT
  TO authenticated
  USING (true);

CREATE POLICY "Authenticated users can insert customers"
  ON customers FOR INSERT
  TO authenticated
  WITH CHECK (true);

CREATE POLICY "Authenticated users can update customers"
  ON customers FOR UPDATE
  TO authenticated
  USING (true)
  WITH CHECK (true);

-- ====================================
-- 商品マスタ
-- ====================================
CREATE TABLE IF NOT EXISTS products (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  product_code text UNIQUE NOT NULL,
  name text NOT NULL,
  category text DEFAULT '',
  standard_price numeric DEFAULT 0,
  unit text DEFAULT '',
  created_at timestamptz DEFAULT now(),
  updated_at timestamptz DEFAULT now()
);

ALTER TABLE products ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Authenticated users can view products"
  ON products FOR SELECT
  TO authenticated
  USING (true);

-- ====================================
-- 単価マスタ
-- ====================================
CREATE TABLE IF NOT EXISTS customer_prices (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  customer_id uuid NOT NULL REFERENCES customers(id) ON DELETE CASCADE,
  product_id uuid NOT NULL REFERENCES products(id) ON DELETE CASCADE,
  price numeric NOT NULL DEFAULT 0,
  effective_date date,
  created_at timestamptz DEFAULT now(),
  updated_at timestamptz DEFAULT now(),
  UNIQUE(customer_id, product_id, effective_date)
);

ALTER TABLE customer_prices ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Authenticated users can view customer prices"
  ON customer_prices FOR SELECT
  TO authenticated
  USING (true);

CREATE POLICY "Authenticated users can manage customer prices"
  ON customer_prices FOR INSERT
  TO authenticated
  WITH CHECK (true);

CREATE POLICY "Authenticated users can update customer prices"
  ON customer_prices FOR UPDATE
  TO authenticated
  USING (true)
  WITH CHECK (true);

CREATE POLICY "Authenticated users can delete customer prices"
  ON customer_prices FOR DELETE
  TO authenticated
  USING (true);

-- ====================================
-- 商談管理
-- ====================================
CREATE TABLE IF NOT EXISTS meetings (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  customer_id uuid NOT NULL REFERENCES customers(id) ON DELETE CASCADE,
  handler_id uuid REFERENCES employees(id),
  schedule_date date NOT NULL,
  actual_date date,
  appointment_type text DEFAULT '',
  purpose text DEFAULT '',
  plan_notes text DEFAULT '',
  result text DEFAULT '',
  result_notes text DEFAULT '',
  delay_reason text DEFAULT '',
  status text DEFAULT 'upcoming',
  created_at timestamptz DEFAULT now(),
  updated_at timestamptz DEFAULT now()
);

ALTER TABLE meetings ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Authenticated users can view meetings"
  ON meetings FOR SELECT
  TO authenticated
  USING (true);

CREATE POLICY "Authenticated users can insert meetings"
  ON meetings FOR INSERT
  TO authenticated
  WITH CHECK (true);

CREATE POLICY "Authenticated users can update meetings"
  ON meetings FOR UPDATE
  TO authenticated
  USING (true)
  WITH CHECK (true);

CREATE POLICY "Authenticated users can delete meetings"
  ON meetings FOR DELETE
  TO authenticated
  USING (true);

-- ====================================
-- 申請管理
-- ====================================
CREATE TABLE IF NOT EXISTS applications (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  application_id text UNIQUE NOT NULL,
  group_id text,
  application_type text NOT NULL,
  target_master text DEFAULT '',
  customer_id uuid REFERENCES customers(id),
  customer_code text,
  customer_name text,
  applicant_id uuid REFERENCES employees(id),
  applicant_name text,
  application_date timestamptz DEFAULT now(),
  status text DEFAULT '申請中',
  stage text DEFAULT '申請中',
  effective_date date,
  reflection_status text DEFAULT '',
  reject_reason text DEFAULT '',
  required_fields jsonb,
  backup_data jsonb,
  price_backup_data jsonb,
  manager_customer_code text,
  manager_group_code text,
  supervisor_approver text,
  supervisor_approved_at timestamptz,
  admin_approver text,
  admin_approved_at timestamptz,
  director_approver text,
  director_approved_at timestamptz,
  ceo_approver text,
  ceo_approved_at timestamptz,
  resubmit_count integer DEFAULT 0,
  original_application_id text,
  created_at timestamptz DEFAULT now(),
  updated_at timestamptz DEFAULT now()
);

ALTER TABLE applications ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Authenticated users can view applications"
  ON applications FOR SELECT
  TO authenticated
  USING (true);

CREATE POLICY "Authenticated users can insert applications"
  ON applications FOR INSERT
  TO authenticated
  WITH CHECK (true);

CREATE POLICY "Authenticated users can update applications"
  ON applications FOR UPDATE
  TO authenticated
  USING (true)
  WITH CHECK (true);

-- ====================================
-- 申請データ_顧客
-- ====================================
CREATE TABLE IF NOT EXISTS application_customer_data (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  application_id uuid NOT NULL REFERENCES applications(id) ON DELETE CASCADE,
  field_name text NOT NULL,
  before_value text DEFAULT '',
  after_value text DEFAULT '',
  created_at timestamptz DEFAULT now()
);

ALTER TABLE application_customer_data ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Authenticated users can view application customer data"
  ON application_customer_data FOR SELECT
  TO authenticated
  USING (true);

CREATE POLICY "Authenticated users can insert application customer data"
  ON application_customer_data FOR INSERT
  TO authenticated
  WITH CHECK (true);

-- ====================================
-- 申請データ_単価
-- ====================================
CREATE TABLE IF NOT EXISTS application_price_data (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  application_id uuid NOT NULL REFERENCES applications(id) ON DELETE CASCADE,
  product_code text NOT NULL,
  product_name text DEFAULT '',
  registration_type text DEFAULT '',
  price_before numeric,
  price_after numeric,
  created_at timestamptz DEFAULT now()
);

ALTER TABLE application_price_data ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Authenticated users can view application price data"
  ON application_price_data FOR SELECT
  TO authenticated
  USING (true);

CREATE POLICY "Authenticated users can insert application price data"
  ON application_price_data FOR INSERT
  TO authenticated
  WITH CHECK (true);

-- ====================================
-- 渋滞金リスト
-- ====================================
CREATE TABLE IF NOT EXISTS jutaikin (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  customer_id uuid NOT NULL REFERENCES customers(id) ON DELETE CASCADE,
  amount numeric DEFAULT 0,
  base_year_month text DEFAULT '',
  created_at timestamptz DEFAULT now()
);

ALTER TABLE jutaikin ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Authenticated users can view jutaikin"
  ON jutaikin FOR SELECT
  TO authenticated
  USING (true);

-- ====================================
-- 売上データ
-- ====================================
CREATE TABLE IF NOT EXISTS customer_sales (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  customer_id uuid NOT NULL REFERENCES customers(id) ON DELETE CASCADE,
  year integer NOT NULL,
  month integer NOT NULL,
  amount numeric DEFAULT 0,
  created_at timestamptz DEFAULT now(),
  UNIQUE(customer_id, year, month)
);

ALTER TABLE customer_sales ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Authenticated users can view customer sales"
  ON customer_sales FOR SELECT
  TO authenticated
  USING (true);

-- ====================================
-- RPA処理用テーブル
-- ====================================
CREATE TABLE IF NOT EXISTS rpa_customer_queue (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  application_id uuid REFERENCES applications(id),
  customer_code text NOT NULL,
  operation_type text NOT NULL,
  data jsonb NOT NULL,
  processed boolean DEFAULT false,
  processed_at timestamptz,
  created_at timestamptz DEFAULT now()
);

ALTER TABLE rpa_customer_queue ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Authenticated users can view rpa customer queue"
  ON rpa_customer_queue FOR SELECT
  TO authenticated
  USING (true);

CREATE POLICY "Authenticated users can insert rpa customer queue"
  ON rpa_customer_queue FOR INSERT
  TO authenticated
  WITH CHECK (true);

CREATE TABLE IF NOT EXISTS rpa_price_queue (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  application_id uuid REFERENCES applications(id),
  customer_code text NOT NULL,
  product_code text NOT NULL,
  operation_type text NOT NULL,
  price numeric,
  effective_date date,
  processed boolean DEFAULT false,
  processed_at timestamptz,
  created_at timestamptz DEFAULT now()
);

ALTER TABLE rpa_price_queue ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Authenticated users can view rpa price queue"
  ON rpa_price_queue FOR SELECT
  TO authenticated
  USING (true);

CREATE POLICY "Authenticated users can insert rpa price queue"
  ON rpa_price_queue FOR INSERT
  TO authenticated
  WITH CHECK (true);

-- ====================================
-- インデックス
-- ====================================
CREATE INDEX IF NOT EXISTS idx_employees_department ON employees(department_id);
CREATE INDEX IF NOT EXISTS idx_employees_login_id ON employees(login_id);
CREATE INDEX IF NOT EXISTS idx_customers_code ON customers(customer_code);
CREATE INDEX IF NOT EXISTS idx_customers_handler ON customers(handler_employee_id);
CREATE INDEX IF NOT EXISTS idx_customer_prices_customer ON customer_prices(customer_id);
CREATE INDEX IF NOT EXISTS idx_customer_prices_product ON customer_prices(product_id);
CREATE INDEX IF NOT EXISTS idx_meetings_customer ON meetings(customer_id);
CREATE INDEX IF NOT EXISTS idx_meetings_handler ON meetings(handler_id);
CREATE INDEX IF NOT EXISTS idx_meetings_schedule ON meetings(schedule_date);
CREATE INDEX IF NOT EXISTS idx_applications_customer ON applications(customer_id);
CREATE INDEX IF NOT EXISTS idx_applications_applicant ON applications(applicant_id);
CREATE INDEX IF NOT EXISTS idx_applications_status ON applications(status);
CREATE INDEX IF NOT EXISTS idx_applications_stage ON applications(stage);

-- database_schema.sql
-- COMPLETE Migrations for Clan Collection Fund fullstack Platform

-- Enable UUID generation
CREATE EXTENSION IF NOT EXISTS "uuid-ossp";

-- ==========================================
-- 1. DEPARTMENTS
-- ==========================================
CREATE TABLE IF NOT EXISTS departments (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    department_name_ar TEXT UNIQUE NOT NULL,
    department_name_en TEXT UNIQUE NOT NULL,
    department_code TEXT UNIQUE NOT NULL,
    is_active BOOLEAN DEFAULT TRUE NOT NULL,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL
);

-- Seed initial departments
INSERT INTO departments (id, department_name_ar, department_name_en, department_code, is_active)
VALUES 
(
    '811be074-ceee-40ba-81df-85e656ad4e81', 
    'اللجنة المالية والتحصيل', 
    'Finance & Collections Committee',
    'DEP-FIN',
    true
),
(
    '52cd6575-cf6d-49b0-bc32-15f1fdfce824', 
    'الأمانة العامة وإدارة الصندوق', 
    'General Secretariat & Administration',
    'DEP-ADM',
    true
)
ON CONFLICT (department_code) DO NOTHING;


-- ==========================================
-- 2. PROFILES (Extends Supabase auth.users)
-- ==========================================
CREATE TABLE IF NOT EXISTS profiles (
    id UUID PRIMARY KEY REFERENCES auth.users(id) ON DELETE CASCADE,
    full_name TEXT NOT NULL,
    email TEXT UNIQUE NOT NULL,
    phone TEXT,
    role TEXT NOT NULL CHECK (role IN (
        'Super Administrator',
        'Administrator',
        'Data Entry User',
        'Reviewer',
        'Approver',
        'Viewer'
    )),
    department_id UUID REFERENCES departments(id) ON DELETE SET NULL,
    job_title TEXT,
    is_active BOOLEAN DEFAULT TRUE NOT NULL,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL
);

-- Enable RLS for Profiles
ALTER TABLE profiles ENABLE ROW LEVEL SECURITY;

-- Select Policies for Profiles
CREATE POLICY "Allow read access to anyone authenticated"
    ON profiles FOR SELECT
    TO authenticated
    USING (true);

-- Insert/Update Policies for Profiles
CREATE POLICY "Super Admins can manage profiles"
    ON profiles FOR ALL
    TO authenticated
    USING (
        (SELECT role FROM profiles WHERE id = auth.uid()) = 'Super Administrator'
    );

CREATE POLICY "Users can update their own contact details"
    ON profiles FOR UPDATE
    TO authenticated
    USING (id = auth.uid())
    WITH CHECK (id = auth.uid() AND role = (SELECT role FROM profiles WHERE id = auth.uid()));


-- ==========================================
-- 3. REPORTS (Consolidates Monthly collections & expenses)
-- ==========================================
CREATE TABLE IF NOT EXISTS reports (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    report_number TEXT UNIQUE NOT NULL,
    report_title TEXT NOT NULL,
    report_type TEXT NOT NULL, 
    department_id UUID REFERENCES departments(id) ON DELETE RESTRICT NOT NULL,
    reporting_period TEXT NOT NULL, -- e.g., "أبريل 2026", "مايو 2026"
    report_date DATE DEFAULT CURRENT_DATE NOT NULL,
    status TEXT NOT NULL CHECK (status IN (
        'Draft',
        'Submitted',
        'Under Review',
        'Returned for Correction',
        'Reviewed',
        'Approved',
        'Rejected',
        'Archived'
    )) DEFAULT 'Draft',
    report_data JSONB NOT NULL, -- Core JSON dump containing inputs (members, payments, expenses)
    rejection_reason TEXT,
    review_comments TEXT,
    created_by UUID REFERENCES auth.users(id) ON DELETE SET NULL,
    reviewed_by UUID REFERENCES auth.users(id) ON DELETE SET NULL,
    approved_by UUID REFERENCES auth.users(id) ON DELETE SET NULL,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL,
    submitted_at TIMESTAMP WITH TIME ZONE,
    reviewed_at TIMESTAMP WITH TIME ZONE,
    approved_at TIMESTAMP WITH TIME ZONE,
    archived_at TIMESTAMP WITH TIME ZONE,
    version_number INTEGER DEFAULT 1 NOT NULL,
    is_deleted BOOLEAN DEFAULT FALSE NOT NULL
);

-- Enable RLS for Reports
ALTER TABLE reports ENABLE ROW LEVEL SECURITY;

-- Read Policy: Members can read reports belonging to their department. Admin and Approvers can access all.
CREATE POLICY "Reports department read policy"
    ON reports FOR SELECT
    TO authenticated
    USING (
        (SELECT role FROM profiles WHERE id = auth.uid()) IN ('Super Administrator', 'Reviewer', 'Approver')
        OR department_id = (SELECT department_id FROM profiles WHERE id = auth.uid())
    );

-- Insert Policy: Data Entry and Superadmin can write drafts
CREATE POLICY "Reports insert drafts policy"
    ON reports FOR INSERT
    TO authenticated
    WITH CHECK (
        (SELECT role FROM profiles WHERE id = auth.uid()) IN ('Super Administrator', 'Data Entry User')
        AND status = 'Draft'
    );

-- Update Policy: Locked if approved (except for admin versions), restrict by role status workflow
CREATE POLICY "Reports edit state policy"
    ON reports FOR UPDATE
    TO authenticated
    USING (
        (SELECT role FROM profiles WHERE id = auth.uid()) = 'Super Administrator'
        OR (
            created_by = auth.uid() 
            AND status IN ('Draft', 'Returned for Correction')
        )
        OR (
            (SELECT role FROM profiles WHERE id = auth.uid()) IN ('Reviewer', 'Approver', 'Administrator')
            AND department_id = (SELECT department_id FROM profiles WHERE id = auth.uid())
        )
    );


-- ==========================================
-- 4. REPORT ATTACHMENTS
-- ==========================================
CREATE TABLE IF NOT EXISTS report_attachments (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    report_id UUID REFERENCES reports(id) ON DELETE CASCADE NOT NULL,
    file_name TEXT NOT NULL,
    file_path TEXT NOT NULL, -- Bucket subpath
    file_type TEXT NOT NULL,
    file_size BIGINT NOT NULL,
    uploaded_by UUID REFERENCES auth.users(id) ON DELETE SET NULL,
    uploaded_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL
);

-- Enable RLS for Attachments
ALTER TABLE report_attachments ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Attachments read policy"
    ON report_attachments FOR SELECT
    TO authenticated
    USING (
        EXISTS (
            SELECT 1 FROM reports 
            WHERE reports.id = report_attachments.report_id
        )
    );

CREATE POLICY "Attachments insert policy"
    ON report_attachments FOR INSERT
    TO authenticated
    WITH CHECK (
        EXISTS (
            SELECT 1 FROM reports 
            WHERE reports.id = report_attachments.report_id 
            AND reports.status IN ('Draft', 'Returned for Correction')
            AND reports.created_by = auth.uid()
        ) OR (SELECT role FROM profiles WHERE id = auth.uid()) = 'Super Administrator'
    );


-- ==========================================
-- 5. REPORT COMMENTS
-- ==========================================
CREATE TABLE IF NOT EXISTS report_comments (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    report_id UUID REFERENCES reports(id) ON DELETE CASCADE NOT NULL,
    comment TEXT NOT NULL,
    comment_type TEXT DEFAULT 'general',
    created_by UUID REFERENCES auth.users(id) ON DELETE SET NULL,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL
);

-- Enable RLS for Comments
ALTER TABLE report_comments ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Comments read policy"
    ON report_comments FOR SELECT
    TO authenticated
    USING (true);

CREATE POLICY "Comments insert policy"
    ON report_comments FOR INSERT
    TO authenticated
    WITH CHECK (auth.uid() = created_by);


-- ==========================================
-- 6. AUDIT LOGS
-- ==========================================
CREATE TABLE IF NOT EXISTS audit_logs (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    user_id UUID REFERENCES auth.users(id) ON DELETE SET NULL,
    action TEXT NOT NULL,
    table_name TEXT NOT NULL,
    record_id UUID,
    old_values JSONB,
    new_values JSONB,
    ip_address TEXT,
    user_agent TEXT,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL
);

-- Enable RLS for Audit Logs: ONLY super admin can read or write
ALTER TABLE audit_logs ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Super admin log policy"
    ON audit_logs FOR SELECT
    TO authenticated
    USING (
        (SELECT role FROM profiles WHERE id = auth.uid()) = 'Super Administrator'
    );


-- ==========================================
-- 7. SYSTEM SETTINGS
-- ==========================================
CREATE TABLE IF NOT EXISTS system_settings (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    setting_key TEXT UNIQUE NOT NULL,
    setting_value JSONB NOT NULL,
    updated_by UUID REFERENCES auth.users(id) ON DELETE SET NULL,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL
);

-- Enable RLS for System Settings
ALTER TABLE system_settings ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Anyone authenticated can select settings"
    ON system_settings FOR SELECT
    TO authenticated
    USING (true);

CREATE POLICY "Admins can update settings"
    ON system_settings FOR UPDATE
    TO authenticated
    USING (
        (SELECT role FROM profiles WHERE id = auth.uid()) in ('Super Administrator', 'Administrator')
    );


-- ==========================================
-- PROCEDURES & TRIGGERS
-- ==========================================

-- Trigger to auto-update update_at column
CREATE OR REPLACE FUNCTION update_modified_column()
RETURNS TRIGGER AS $$
BEGIN
    NEW.updated_at = now();
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER update_profiles_modtime
    BEFORE UPDATE ON profiles
    FOR EACH ROW EXECUTE PROCEDURE update_modified_column();

CREATE TRIGGER update_reports_modtime
    BEFORE UPDATE ON reports
    FOR EACH ROW EXECUTE PROCEDURE update_modified_column();

-- Auto Sync new users into public.profiles
CREATE OR REPLACE FUNCTION public.handle_new_user()
RETURNS trigger AS $$
BEGIN
  INSERT INTO public.profiles (id, full_name, email, role, is_active)
  VALUES (
    new.id,
    coalesce(new.raw_user_meta_data->>'full_name', 'مستخدم جديد'),
    new.email,
    coalesce(new.raw_user_meta_data->>'role', 'Viewer'),
    true
  );
  RETURN new;
END;
$$ LANGUAGE plpgsql SECURITY DEFINER;

CREATE OR REPLACE TRIGGER on_auth_user_created
  AFTER INSERT ON auth.users
  FOR EACH ROW EXECUTE PROCEDURE public.handle_new_user();

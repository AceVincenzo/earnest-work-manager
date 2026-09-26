-- =========================================================
-- EARNEST - SUPABASE POSTGRESQL DATABASE SCHEMA & RLS POLICIES
-- =========================================================

-- Enable UUID extension
CREATE EXTENSION IF NOT EXISTS "uuid-ossp";

-- 1. PROFILES TABLE
CREATE TABLE IF NOT EXISTS public.profiles (
    id UUID PRIMARY KEY REFERENCES auth.users(id) ON DELETE CASCADE,
    name TEXT NOT NULL,
    avatar TEXT DEFAULT 'M',
    is_admin BOOLEAN DEFAULT false,
    pin_code VARCHAR(4),
    created_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL
);

-- 2. EMPLOYERS TABLE
CREATE TABLE IF NOT EXISTS public.employers (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    user_id UUID NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
    name TEXT NOT NULL,
    location TEXT,
    default_tax_rate NUMERIC(4,2) DEFAULT 0.32,
    notes TEXT,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL
);

-- 3. ROLES TABLE
CREATE TABLE IF NOT EXISTS public.roles (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    employer_id UUID NOT NULL REFERENCES public.employers(id) ON DELETE CASCADE,
    user_id UUID NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
    role_name TEXT NOT NULL,
    hourly_rate NUMERIC(10,2) NOT NULL,
    rate_type TEXT DEFAULT 'gross',
    default_break_mins INT DEFAULT 30,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL
);

-- 4. SHIFTS TABLE
CREATE TABLE IF NOT EXISTS public.shifts (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    user_id UUID NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
    employer_id UUID REFERENCES public.employers(id) ON DELETE SET NULL,
    employer_name TEXT NOT NULL,
    role_name TEXT NOT NULL,
    date DATE NOT NULL,
    start_time TIME NOT NULL,
    end_time TIME NOT NULL,
    unpaid_break_mins INT DEFAULT 0,
    paid_hours NUMERIC(6,2) NOT NULL,
    hourly_rate NUMERIC(10,2) NOT NULL,
    gross_pay NUMERIC(10,2) NOT NULL,
    ob_pay NUMERIC(10,2) DEFAULT 0.00,
    tips NUMERIC(10,2) DEFAULT 0.00,
    expenses NUMERIC(10,2) DEFAULT 0.00,
    net_pay NUMERIC(10,2) NOT NULL,
    status TEXT DEFAULT 'worked' CHECK (status IN ('worked', 'scheduled', 'draft', 'attention')),
    notes TEXT,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL
);

-- 5. PAYSLIPS TABLE
CREATE TABLE IF NOT EXISTS public.payslips (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    user_id UUID NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
    employer_id UUID REFERENCES public.employers(id) ON DELETE SET NULL,
    employer_name TEXT NOT NULL,
    pay_period_start DATE NOT NULL,
    pay_period_end DATE NOT NULL,
    actual_gross NUMERIC(10,2) NOT NULL,
    actual_net NUMERIC(10,2) NOT NULL,
    tax_withheld NUMERIC(10,2) DEFAULT 0.00,
    pension NUMERIC(10,2) DEFAULT 0.00,
    vacation_pay NUMERIC(10,2) DEFAULT 0.00,
    deductions NUMERIC(10,2) DEFAULT 0.00,
    notes TEXT,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL
);

-- 6. USER GOALS TABLE
CREATE TABLE IF NOT EXISTS public.user_goals (
    user_id UUID PRIMARY KEY REFERENCES auth.users(id) ON DELETE CASCADE,
    hours_target NUMERIC(6,2) DEFAULT 100.0,
    net_target NUMERIC(10,2) DEFAULT 15000.0,
    shifts_target INT DEFAULT 20,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL
);

-- =========================================================
-- ROW LEVEL SECURITY (RLS) POLICIES
-- =========================================================

ALTER TABLE public.profiles ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.employers ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.roles ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.shifts ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.payslips ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.user_goals ENABLE ROW LEVEL SECURITY;

-- PROFILES POLICIES
CREATE POLICY "Users can view own profile" ON public.profiles FOR SELECT USING (auth.uid() = id);
CREATE POLICY "Users can update own profile" ON public.profiles FOR UPDATE USING (auth.uid() = id);
CREATE POLICY "Users can insert own profile" ON public.profiles FOR INSERT WITH CHECK (auth.uid() = id);

-- EMPLOYERS POLICIES
CREATE POLICY "Users can manage own employers" ON public.employers FOR ALL USING (auth.uid() = user_id);

-- ROLES POLICIES
CREATE POLICY "Users can manage own roles" ON public.roles FOR ALL USING (auth.uid() = user_id);

-- SHIFTS POLICIES
CREATE POLICY "Users can manage own shifts" ON public.shifts FOR ALL USING (auth.uid() = user_id);

-- PAYSLIPS POLICIES
CREATE POLICY "Users can manage own payslips" ON public.payslips FOR ALL USING (auth.uid() = user_id);

-- GOALS POLICIES
CREATE POLICY "Users can manage own goals" ON public.user_goals FOR ALL USING (auth.uid() = user_id);

-- Trigger to auto-create profile on Auth signup
CREATE OR REPLACE FUNCTION public.handle_new_user()
RETURNS TRIGGER AS $$
BEGIN
  INSERT INTO public.profiles (id, name, avatar, is_admin)
  VALUES (new.id, COALESCE(new.raw_user_meta_data->>'full_name', new.email), UPPER(LEFT(COALESCE(new.raw_user_meta_data->>'full_name', new.email), 1)), true);

  INSERT INTO public.user_goals (user_id, hours_target, net_target, shifts_target)
  VALUES (new.id, 100.0, 15000.0, 20);

  RETURN NEW;
END;
$$ LANGUAGE plpgsql SECURITY DEFINER;

CREATE OR REPLACE TRIGGER on_auth_user_created
  AFTER INSERT ON auth.users
  FOR EACH ROW EXECUTE FUNCTION public.handle_new_user();

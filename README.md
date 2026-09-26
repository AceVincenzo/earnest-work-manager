# 💼 Earnest — Personal Work & Payroll Management System

Earnest is a web application designed for restaurant, hotel, shift, and multi-job workers to track hours, calculate gross/net wages, reconcile payslips, manage shift schedules, and analyze personal earnings.

---

## 🌟 Key Features

- **🔥 Quick Shift Logger + Save & Duplicate**: Log shifts with OB, tips, break rules, and expenses. Duplicate shift entries with 1 click.
- **📅 Interactive Shift Calendar**: Color-coded view of Completed (🟢), Scheduled (🔵), Draft (🟡), and Attention (🔴) shifts.
- **🗓️ Scheduled Shifts Workflow**: Life-cycle tracking from `Scheduled` → `Worked` → `Paid` with a 1-tap **"Mark as Worked ✓"** button.
- **💰 Expected Paycheck Engine**: Itemized breakdown of Base wages, OB evening/weekend pay, Overtime, Tips, Taxes, Expenses, and Net salary.
- **🧾 Payslip Scanner & Reconciliation**: Compare calculated shift earnings against official payslips with a **"Why is my salary different?"** variance analyzer.
- **📊 Analytics & "What If?" Simulator**: Calculate effective hourly net earnings (`Net ÷ Worked Hours`) and simulate extra shift hours or rate increases.
- **🎯 Monthly Goal Tracking**: Set targets for hours, net income, and total shifts with progress bars.
- **🌐 Supabase PostgreSQL & Auth**: Persistent cloud database with Row Level Security (RLS) policies and local storage migration.
- **💻 Local PC Python Server**: Runs 100% offline via `python server.py` using persistent SQLite.
- **📱 PWA & Apple iCal Sync**: Progressive Web App support ("Add to Home Screen") and 1-tap iOS Safari Apple Calendar sync.

---

## 🚀 Quick Setup & Local Execution

### Option A: Local PC (Python Server)
1. Ensure Python 3 is installed.
2. Run the server:
   ```bash
   cd app
   python server.py
   ```
3. Open `http://127.0.0.1:5000` or `http://localhost:5000` in your browser.

### Option B: Standalone Web App / Netlify / Vercel
1. Open `public/index.html` directly in any web browser.
2. Deploy the `public` directory to Vercel or Netlify.

---

## 🌐 Supabase PostgreSQL Setup

1. Create a free project at [Supabase.com](https://supabase.com).
2. Execute the `supabase_schema.sql` script in your Supabase SQL Editor.
3. In Earnest, navigate to **Settings & Sync** → **Supabase Cloud Database & Auth**.
4. Paste your **Supabase URL** & **Anon API Key** and click **Connect Supabase**.
5. Sign in or Sign up, then click **Migrate Local Data to Cloud** to sync your history permanently.

---

## 📄 License

MIT License. Designed with DM Serif Display & Plus Jakarta Sans.

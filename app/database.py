import sqlite3
import os
import datetime
from calculator import calculate_shift

DB_PATH = os.path.join(os.path.dirname(__file__), "work_manager.db")

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db(reset_demo=False):
    if reset_demo and os.path.exists(DB_PATH):
        try:
            os.remove(DB_PATH)
        except Exception:
            pass

    conn = get_db()
    cursor = conn.cursor()

    # 1. Employers Table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS employers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE,
        default_tax_rate REAL DEFAULT 0.32,
        notes TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """)

    # 2. Roles Table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS roles (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        employer_id INTEGER NOT NULL,
        role_name TEXT NOT NULL,
        hourly_rate REAL NOT NULL,
        rate_type TEXT DEFAULT 'gross', -- 'gross' or 'net'
        default_break_mins INTEGER DEFAULT 30,
        FOREIGN KEY(employer_id) REFERENCES employers(id) ON DELETE CASCADE
    );
    """)

    # 3. Shifts Table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS shifts (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        employer_id INTEGER NOT NULL,
        role_id INTEGER,
        date TEXT NOT NULL,          -- YYYY-MM-DD
        start_time TEXT NOT NULL,    -- HH:MM
        end_time TEXT NOT NULL,      -- HH:MM
        unpaid_break_mins INTEGER DEFAULT 30,
        hourly_rate REAL NOT NULL,   -- Historical rate snapshot
        rate_type TEXT DEFAULT 'gross',
        tax_rate REAL DEFAULT 0.32,
        ob_pay REAL DEFAULT 0.0,
        overtime_pay REAL DEFAULT 0.0,
        tips REAL DEFAULT 0.0,
        bonus REAL DEFAULT 0.0,
        expenses REAL DEFAULT 0.0,
        paid_hours REAL NOT NULL,
        gross_pay REAL NOT NULL,
        tax_amount REAL NOT NULL,
        net_pay REAL NOT NULL,
        notes TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(employer_id) REFERENCES employers(id),
        FOREIGN KEY(role_id) REFERENCES roles(id)
    );
    """)

    # 4. Payslips Table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS payslips (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        employer_id INTEGER NOT NULL,
        pay_period_start TEXT NOT NULL, -- YYYY-MM-DD
        pay_period_end TEXT NOT NULL,   -- YYYY-MM-DD
        payslip_date TEXT NOT NULL,     -- YYYY-MM-DD
        actual_gross REAL NOT NULL,
        tax_withheld REAL NOT NULL,
        actual_net REAL NOT NULL,
        other_deductions REAL DEFAULT 0.0,
        notes TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(employer_id) REFERENCES employers(id)
    );
    """)

    conn.commit()

    # Check if data exists; if not, seed default demo data!
    cursor.execute("SELECT COUNT(*) as cnt FROM employers")
    if cursor.fetchone()['cnt'] == 0:
        seed_demo_data(conn)

    conn.close()

def seed_demo_data(conn):
    cursor = conn.cursor()
    
    # Add Employers
    employers_data = [
        ("Agra Tandoori", 0.32, "Primary evening job. Flat tax or employer handles tax."),
        ("Grand Hôtel", 0.32, "Event server & banquet runner. Total salary 153.64 kr/hr."),
        ("Backup Job", 0.30, "Flexible secondary job.")
    ]
    
    emp_ids = {}
    for name, tax, notes in employers_data:
        cursor.execute("INSERT INTO employers (name, default_tax_rate, notes) VALUES (?, ?, ?)", (name, tax, notes))
        emp_ids[name] = cursor.lastrowid

    # Add Roles
    roles_data = [
        (emp_ids["Agra Tandoori"], "Kitchen & Server", 95.00, "gross", 30),
        (emp_ids["Grand Hôtel"], "Event Runner", 153.64, "gross", 30),
        (emp_ids["Grand Hôtel"], "Banquet Captain", 165.00, "gross", 30),
        (emp_ids["Backup Job"], "Store Assistant", 120.00, "gross", 30)
    ]
    
    role_ids = {}
    for emp_id, r_name, rate, r_type, break_m in roles_data:
        cursor.execute("INSERT INTO roles (employer_id, role_name, hourly_rate, rate_type, default_break_mins) VALUES (?, ?, ?, ?, ?)",
                       (emp_id, r_name, rate, r_type, break_m))
        role_ids[f"{emp_id}_{r_name}"] = cursor.lastrowid

    # Sample Shifts
    demo_shifts = [
        # Date, Employer Name, Role Name, Start, End, Break, OB, Tips, Overtime, Notes
        ("2026-09-01", "Agra Tandoori", "Kitchen & Server", "11:00", "19:00", 30, 0, 150, 0, "Lunch & early dinner"),
        ("2026-09-03", "Agra Tandoori", "Kitchen & Server", "16:00", "23:00", 30, 100, 220, 0, "Busy evening shift"),
        ("2026-09-05", "Grand Hôtel", "Event Runner", "15:00", "23:30", 45, 150, 300, 0, "Banquet gala dinner"),
        ("2026-09-08", "Grand Hôtel", "Event Runner", "07:00", "15:30", 30, 0, 100, 0, "Morning conference prep"),
        ("2026-09-11", "Agra Tandoori", "Kitchen & Server", "16:00", "23:30", 30, 0, 180, 0, "Friday night service"),
        ("2026-09-14", "Backup Job", "Store Assistant", "09:00", "17:00", 30, 0, 0, 0, "Inventory shift"),
        ("2026-09-17", "Grand Hôtel", "Banquet Captain", "14:00", "22:00", 30, 200, 250, 150, "Lead event shift"),
        ("2026-09-20", "Agra Tandoori", "Kitchen & Server", "12:00", "22:00", 60, 100, 350, 0, "Full Sunday shift"),
        ("2026-09-23", "Grand Hôtel", "Event Runner", "18:00", "02:00", 30, 250, 200, 0, "Overnight wedding event"),
        ("2026-09-25", "Agra Tandoori", "Kitchen & Server", "16:00", "23:00", 30, 0, 190, 0, "Friday shift"),
    ]

    for d_str, emp_name, r_name, s_t, e_t, b_m, ob, tips, ot, note in demo_shifts:
        e_id = emp_ids[emp_name]
        r_id = role_ids[f"{e_id}_{r_name}"]
        
        # Fetch rate & tax
        cursor.execute("SELECT default_tax_rate FROM employers WHERE id = ?", (e_id,))
        tax_rate = cursor.fetchone()['default_tax_rate']

        cursor.execute("SELECT hourly_rate, rate_type FROM roles WHERE id = ?", (r_id,))
        r_row = cursor.fetchone()
        rate = r_row['hourly_rate']
        rate_type = r_row['rate_type']

        calc = calculate_shift(d_str, s_t, e_t, break_mins=b_m, hourly_rate=rate,
                               rate_type=rate_type, tax_rate=tax_rate,
                               ob_pay=ob, overtime_pay=ot, tips=tips)

        cursor.execute("""
        INSERT INTO shifts (employer_id, role_id, date, start_time, end_time, unpaid_break_mins,
                            hourly_rate, rate_type, tax_rate, ob_pay, overtime_pay, tips, bonus, expenses,
                            paid_hours, gross_pay, tax_amount, net_pay, notes)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (e_id, r_id, d_str, s_t, e_t, b_m, rate, rate_type, tax_rate, ob, ot, tips, 0.0, 0.0,
              calc['paid_hours'], calc['total_gross'], calc['estimated_tax'], calc['estimated_net'], note))

    # No default payslips added, payslips table left clean for real user entry
    conn.commit()
    print("Demo data successfully seeded!")

if __name__ == "__main__":
    init_db(reset_demo=True)

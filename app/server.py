import os
import sqlite3
import datetime
import io
import json
from flask import Flask, render_template, request, jsonify, send_file, Response
import openpyxl

from database import get_db, init_db
from calculator import calculate_shift

app = Flask(__name__, template_folder="templates")

# Ensure DB initialized
init_db(reset_demo=False)

@app.route("/")
def index():
    return render_template("index.html")

# ---------------------------------------------------------
# DASHBOARD API
# ---------------------------------------------------------
@app.route("/api/dashboard", methods=["GET"])
def get_dashboard():
    range_filter = request.args.get("filter", "month") # today, week, month, year, all, custom
    custom_start = request.args.get("start")
    custom_end = request.args.get("end")

    today = datetime.date.today()
    
    # Calculate date bounds
    if range_filter == "today":
        start_date = today.strftime("%Y-%m-%d")
        end_date = today.strftime("%Y-%m-%d")
    elif range_filter == "week":
        start_date = (today - datetime.timedelta(days=today.weekday())).strftime("%Y-%m-%d")
        end_date = (today + datetime.timedelta(days=6-today.weekday())).strftime("%Y-%m-%d")
    elif range_filter == "month":
        start_date = today.replace(day=1).strftime("%Y-%m-%d")
        # Next month start minus 1 day
        next_m = today.replace(day=28) + datetime.timedelta(days=4)
        end_date = (next_m - datetime.timedelta(days=next_m.day)).strftime("%Y-%m-%d")
    elif range_filter == "year":
        start_date = f"{today.year}-01-01"
        end_date = f"{today.year}-12-31"
    elif range_filter == "custom" and custom_start and custom_end:
        start_date = custom_start
        end_date = custom_end
    else:
        start_date = "1970-01-01"
        end_date = "2099-12-31"

    conn = get_db()
    cursor = conn.cursor()

    # KPI summary
    cursor.execute("""
    SELECT 
        COUNT(*) as total_shifts,
        COALESCE(SUM(paid_hours), 0) as total_hours,
        COALESCE(SUM(gross_pay), 0) as total_gross,
        COALESCE(SUM(tax_amount), 0) as total_tax,
        COALESCE(SUM(net_pay), 0) as total_net,
        COALESCE(SUM(tips + ob_pay + overtime_pay + bonus - expenses), 0) as total_supplements
    FROM shifts
    WHERE date >= ? AND date <= ?
    """, (start_date, end_date))
    kpi_row = cursor.fetchone()

    total_shifts = kpi_row['total_shifts']
    total_hours = round(kpi_row['total_hours'], 2)
    total_gross = round(kpi_row['total_gross'], 2)
    total_tax = round(kpi_row['total_tax'], 2)
    total_net = round(kpi_row['total_net'], 2)
    total_supplements = round(kpi_row['total_supplements'], 2)

    avg_hourly_net = round(total_net / total_hours, 2) if total_hours > 0 else 0.0

    # Employer Breakdown
    cursor.execute("""
    SELECT 
        e.name as employer_name,
        COALESCE(SUM(s.paid_hours), 0) as hours,
        COALESCE(SUM(s.gross_pay), 0) as gross,
        COALESCE(SUM(s.net_pay), 0) as net,
        COUNT(s.id) as shift_count
    FROM employers e
    LEFT JOIN shifts s ON e.id = s.employer_id AND s.date >= ? AND s.date <= ?
    GROUP BY e.id
    ORDER BY net DESC
    """, (start_date, end_date))
    employer_breakdown = [dict(row) for row in cursor.fetchall()]

    # Monthly Trend (Last 12 months)
    cursor.execute("""
    SELECT 
        strftime('%Y-%m', date) as month_key,
        COALESCE(SUM(paid_hours), 0) as hours,
        COALESCE(SUM(gross_pay), 0) as gross,
        COALESCE(SUM(net_pay), 0) as net
    FROM shifts
    GROUP BY strftime('%Y-%m', date)
    ORDER BY month_key ASC
    """)
    monthly_trend = [dict(row) for row in cursor.fetchall()]

    # Weekday Distribution
    cursor.execute("""
    SELECT 
        case strftime('%w', date)
            when '0' then 'Sunday'
            when '1' then 'Monday'
            when '2' then 'Tuesday'
            when '3' then 'Wednesday'
            when '4' then 'Thursday'
            when '5' then 'Friday'
            when '6' then 'Saturday'
        end as weekday,
        COALESCE(SUM(paid_hours), 0) as hours,
        COALESCE(SUM(net_pay), 0) as net
    FROM shifts
    WHERE date >= ? AND date <= ?
    GROUP BY strftime('%w', date)
    ORDER BY strftime('%w', date) ASC
    """, (start_date, end_date))
    weekday_distribution = [dict(row) for row in cursor.fetchall()]

    # Current vs Previous Month Comparison
    cur_m_start = today.replace(day=1).strftime("%Y-%m-%d")
    prev_m_end = (today.replace(day=1) - datetime.timedelta(days=1)).strftime("%Y-%m-%d")
    prev_m_start = (today.replace(day=1) - datetime.timedelta(days=1)).replace(day=1).strftime("%Y-%m-%d")

    cursor.execute("SELECT COALESCE(SUM(net_pay), 0) as net, COALESCE(SUM(paid_hours), 0) as hours FROM shifts WHERE date >= ? AND date <= ?", (cur_m_start, today.strftime("%Y-%m-%d")))
    cur_m_data = cursor.fetchone()

    cursor.execute("SELECT COALESCE(SUM(net_pay), 0) as net, COALESCE(SUM(paid_hours), 0) as hours FROM shifts WHERE date >= ? AND date <= ?", (prev_m_start, prev_m_end))
    prev_m_data = cursor.fetchone()

    conn.close()

    return jsonify({
        "kpis": {
            "total_shifts": total_shifts,
            "total_hours": total_hours,
            "total_gross": total_gross,
            "total_tax": total_tax,
            "total_net": total_net,
            "total_supplements": total_supplements,
            "avg_hourly_net": avg_hourly_net
        },
        "employer_breakdown": employer_breakdown,
        "monthly_trend": monthly_trend,
        "weekday_distribution": weekday_distribution,
        "month_comparison": {
            "current_net": round(cur_m_data['net'], 2),
            "current_hours": round(cur_m_data['hours'], 2),
            "prev_net": round(prev_m_data['net'], 2),
            "prev_hours": round(prev_m_data['hours'], 2)
        }
    })

# ---------------------------------------------------------
# SHIFTS API
# ---------------------------------------------------------
@app.route("/api/shifts", methods=["GET"])
def get_shifts():
    employer_id = request.args.get("employer_id")
    search = request.args.get("search", "")

    conn = get_db()
    cursor = conn.cursor()

    query = """
    SELECT 
        s.*,
        e.name as employer_name,
        r.role_name
    FROM shifts s
    JOIN employers e ON s.employer_id = e.id
    LEFT JOIN roles r ON s.role_id = r.id
    WHERE 1=1
    """
    params = []

    if employer_id:
        query += " AND s.employer_id = ?"
        params.append(employer_id)

    if search:
        query += " AND (s.notes LIKE ? OR e.name LIKE ? OR r.role_name LIKE ?)"
        s_term = f"%{search}%"
        params.extend([s_term, s_term, s_term])

    query += " ORDER BY s.date DESC, s.start_time DESC"

    cursor.execute(query, params)
    shifts = [dict(row) for row in cursor.fetchall()]
    conn.close()

    return jsonify(shifts)

@app.route("/api/shifts", methods=["POST"])
def add_shift():
    data = request.json
    try:
        emp_id = data["employer_id"]
        role_id = data.get("role_id")
        date_str = data["date"]
        start_time = data["start_time"]
        end_time = data["end_time"]
        break_mins = int(data.get("unpaid_break_mins", 30))
        
        rate = float(data["hourly_rate"])
        rate_type = data.get("rate_type", "gross")
        tax_rate = float(data.get("tax_rate", 0.32))

        ob_pay = float(data.get("ob_pay", 0.0))
        overtime_pay = float(data.get("overtime_pay", 0.0))
        tips = float(data.get("tips", 0.0))
        bonus = float(data.get("bonus", 0.0))
        expenses = float(data.get("expenses", 0.0))
        notes = data.get("notes", "")

        calc = calculate_shift(date_str, start_time, end_time, break_mins=break_mins,
                               hourly_rate=rate, rate_type=rate_type, tax_rate=tax_rate,
                               ob_pay=ob_pay, overtime_pay=overtime_pay, tips=tips, bonus=bonus, expenses=expenses)

        conn = get_db()
        cursor = conn.cursor()
        cursor.execute("""
        INSERT INTO shifts (employer_id, role_id, date, start_time, end_time, unpaid_break_mins,
                            hourly_rate, rate_type, tax_rate, ob_pay, overtime_pay, tips, bonus, expenses,
                            paid_hours, gross_pay, tax_amount, net_pay, notes)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (emp_id, role_id, date_str, start_time, end_time, break_mins,
              rate, rate_type, tax_rate, ob_pay, overtime_pay, tips, bonus, expenses,
              calc["paid_hours"], calc["total_gross"], calc["estimated_tax"], calc["estimated_net"], notes))
        conn.commit()
        shift_id = cursor.lastrowid
        conn.close()

        return jsonify({"success": True, "shift_id": shift_id, "calc": calc})

    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 400

@app.route("/api/shifts/<int:shift_id>", methods=["DELETE"])
def delete_shift(shift_id):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM shifts WHERE id = ?", (shift_id,))
    conn.commit()
    conn.close()
    return jsonify({"success": True})

@app.route("/api/shifts/duplicate/<int:shift_id>", methods=["POST"])
def duplicate_shift(shift_id):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM shifts WHERE id = ?", (shift_id,))
    shift = cursor.fetchone()
    if not shift:
        conn.close()
        return jsonify({"success": False, "error": "Shift not found"}), 404

    s = dict(shift)
    target_date = request.json.get("date", datetime.date.today().strftime("%Y-%m-%d"))

    cursor.execute("""
    INSERT INTO shifts (employer_id, role_id, date, start_time, end_time, unpaid_break_mins,
                        hourly_rate, rate_type, tax_rate, ob_pay, overtime_pay, tips, bonus, expenses,
                        paid_hours, gross_pay, tax_amount, net_pay, notes)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (s['employer_id'], s['role_id'], target_date, s['start_time'], s['end_time'], s['unpaid_break_mins'],
          s['hourly_rate'], s['rate_type'], s['tax_rate'], s['ob_pay'], s['overtime_pay'], s['tips'], s['bonus'], s['expenses'],
          s['paid_hours'], s['gross_pay'], s['tax_amount'], s['net_pay'], f"Duplicated from shift #{shift_id}"))
    conn.commit()
    new_id = cursor.lastrowid
    conn.close()

    return jsonify({"success": True, "new_shift_id": new_id})

# ---------------------------------------------------------
# EMPLOYERS & ROLES API
# ---------------------------------------------------------
@app.route("/api/employers", methods=["GET"])
def get_employers():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM employers ORDER BY name ASC")
    employers = [dict(row) for row in cursor.fetchall()]

    for emp in employers:
        cursor.execute("SELECT * FROM roles WHERE employer_id = ? ORDER BY role_name ASC", (emp['id'],))
        emp['roles'] = [dict(r) for r in cursor.fetchall()]

    conn.close()
    return jsonify(employers)

@app.route("/api/employers", methods=["POST"])
def add_employer():
    data = request.json
    name = data.get("name")
    tax_rate = float(data.get("default_tax_rate", 0.32))
    notes = data.get("notes", "")

    if not name:
        return jsonify({"success": False, "error": "Employer name is required"}), 400

    conn = get_db()
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO employers (name, default_tax_rate, notes) VALUES (?, ?, ?)", (name, tax_rate, notes))
        emp_id = cursor.lastrowid
        
        # Optionally add default role if provided
        roles = data.get("roles", [])
        for r in roles:
            cursor.execute("INSERT INTO roles (employer_id, role_name, hourly_rate, rate_type, default_break_mins) VALUES (?, ?, ?, ?, ?)",
                           (emp_id, r.get("role_name", "General Shift"), float(r.get("hourly_rate", 100)), r.get("rate_type", "gross"), int(r.get("default_break_mins", 30))))
        
        conn.commit()
        conn.close()
        return jsonify({"success": True, "employer_id": emp_id})
    except sqlite3.IntegrityError:
        conn.close()
        return jsonify({"success": False, "error": "Employer name already exists"}), 400

# ---------------------------------------------------------
# PAYSLIPS API & RECONCILIATION
# ---------------------------------------------------------
@app.route("/api/payslips", methods=["GET"])
def get_payslips():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
    SELECT 
        p.*,
        e.name as employer_name,
        (SELECT COALESCE(SUM(gross_pay), 0) FROM shifts WHERE employer_id = p.employer_id AND date >= p.pay_period_start AND date <= p.pay_period_end) as est_gross,
        (SELECT COALESCE(SUM(tax_amount), 0) FROM shifts WHERE employer_id = p.employer_id AND date >= p.pay_period_start AND date <= p.pay_period_end) as est_tax,
        (SELECT COALESCE(SUM(net_pay), 0) FROM shifts WHERE employer_id = p.employer_id AND date >= p.pay_period_start AND date <= p.pay_period_end) as est_net
    FROM payslips p
    JOIN employers e ON p.employer_id = e.id
    ORDER BY p.payslip_date DESC
    """)
    payslips = [dict(row) for row in cursor.fetchall()]

    for p in payslips:
        p["variance_net"] = round(p["actual_net"] - p["est_net"], 2)
        p["variance_gross"] = round(p["actual_gross"] - p["est_gross"], 2)

    conn.close()
    return jsonify(payslips)

@app.route("/api/payslips", methods=["POST"])
def add_payslip():
    data = request.json
    try:
        emp_id = data["employer_id"]
        p_start = data["pay_period_start"]
        p_end = data["pay_period_end"]
        p_date = data["payslip_date"]
        a_gross = float(data["actual_gross"])
        tax_w = float(data["tax_withheld"])
        a_net = float(data["actual_net"])
        other_ded = float(data.get("other_deductions", 0.0))
        notes = data.get("notes", "")

        conn = get_db()
        cursor = conn.cursor()
        cursor.execute("""
        INSERT INTO payslips (employer_id, pay_period_start, pay_period_end, payslip_date, actual_gross, tax_withheld, actual_net, other_deductions, notes)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (emp_id, p_start, p_end, p_date, a_gross, tax_w, a_net, other_ded, notes))
        conn.commit()
        p_id = cursor.lastrowid
        conn.close()

        return jsonify({"success": True, "payslip_id": p_id})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 400

# ---------------------------------------------------------
# EXPORT & DEMO RESET
# ---------------------------------------------------------
@app.route("/api/export/excel", methods=["GET"])
def export_excel():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
    SELECT s.date, e.name as employer, r.role_name, s.start_time, s.end_time, s.unpaid_break_mins,
           s.hourly_rate, s.tax_rate, s.paid_hours, s.gross_pay, s.tax_amount, s.net_pay, s.notes
    FROM shifts s
    JOIN employers e ON s.employer_id = e.id
    LEFT JOIN roles r ON s.role_id = r.id
    ORDER BY s.date DESC
    """)
    rows = cursor.fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Shifts Log"

    headers = ["Date", "Employer", "Role", "Start", "End", "Break (m)", "Rate (kr)", "Tax %", "Paid Hours", "Gross (kr)", "Tax (kr)", "Net (kr)", "Notes"]
    ws.append(headers)

    for row in rows:
        ws.append(list(row))

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True, download_name="Work_Shifts_Export.xlsx")

@app.route("/api/reset-demo", methods=["POST"])
def reset_demo():
    init_db(reset_demo=True)
    return jsonify({"success": True, "message": "Demo data successfully reset!"})

# ---------------------------------------------------------
# APPLE CALENDAR SUBSCRIPTION (.ics / webcal) API
# ---------------------------------------------------------
@app.route("/api/calendar.ics")
@app.route("/shifts.ics")
def get_calendar_ics():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
    SELECT s.*, e.name as employer_name, r.role_name
    FROM shifts s
    JOIN employers e ON s.employer_id = e.id
    LEFT JOIN roles r ON s.role_id = r.id
    ORDER BY s.date ASC, s.start_time ASC
    """)
    shifts = [dict(row) for row in cursor.fetchall()]
    conn.close()

    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//Work & Salary Manager//EN",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        "X-WR-CALNAME:Work Shifts & Wages"
    ]

    for s in shifts:
        s_date = datetime.datetime.strptime(s["date"], "%Y-%m-%d").date()
        s_time = datetime.datetime.strptime(s["start_time"], "%H:%M").time()
        e_time = datetime.datetime.strptime(s["end_time"], "%H:%M").time()

        start_dt = datetime.datetime.combine(s_date, s_time)
        if e_time <= s_time:
            end_dt = datetime.datetime.combine(s_date + datetime.timedelta(days=1), e_time)
        else:
            end_dt = datetime.datetime.combine(s_date, e_time)

        dtstart_str = start_dt.strftime("%Y%m%dT%H%M%S")
        dtend_str = end_dt.strftime("%Y%m%dT%H%M%S")
        summary = f"Work @ {s['employer_name']} ({s['role_name'] or 'Shift'})"
        desc = f"Paid Hours: {s['paid_hours']}h | Est Net: {s['net_pay']} kr (Gross: {s['gross_pay']} kr) | Notes: {s['notes'] or 'None'}"

        lines.extend([
            "BEGIN:VEVENT",
            f"UID:shift-{s['id']}@workmanager.local",
            f"DTSTART:{dtstart_str}",
            f"DTEND:{dtend_str}",
            f"SUMMARY:{summary}",
            f"DESCRIPTION:{desc}",
            "STATUS:CONFIRMED",
            "END:VEVENT"
        ])

    lines.append("END:VCALENDAR")
    ics_content = "\r\n".join(lines)

    return Response(ics_content, mimetype="text/calendar", headers={
        "Content-Disposition": "inline; filename=shifts.ics",
        "Cache-Control": "no-cache"
    })

# ---------------------------------------------------------
# PROFILES & DEVICE SYNC API
# ---------------------------------------------------------
PROFILES_FILE = os.path.join(os.path.dirname(__file__), "profiles_data.json")

@app.route("/api/profiles", methods=["GET"])
def get_profiles_sync():
    if os.path.exists(PROFILES_FILE):
        try:
            with open(PROFILES_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return jsonify({"profiles": data})
        except Exception:
            pass
    return jsonify({"profiles": []})

@app.route("/api/profiles", methods=["POST"])
def save_profiles_sync():
    try:
        data = request.json
        profiles_list = data.get("profiles", [])
        with open(PROFILES_FILE, "w", encoding="utf-8") as f:
            json.dump(profiles_list, f, indent=2)
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 400

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)

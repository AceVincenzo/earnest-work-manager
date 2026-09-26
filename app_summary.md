# 🚀 Work & Salary Manager Web Application

Your personal **Work Hours & Wage Manager** has been upgraded from a basic spreadsheet into a modern, mobile-friendly **Full-Stack Web Application**!

---

## 🌟 Key Application Features

### 1. ⚡ Quick Add Shift & Overnight Shift Support
- **Overnight Shifts**: Supports shifts crossing midnight (e.g. `18:00` to `02:00`). Automatically detects the end date on the next calendar day and calculates exact elapsed hours minus break without negative values or calculation errors.
- **Supplements & Tips**: Log base hourly rate alongside **OB (evening/weekend) supplements**, **Overtime**, **Tips**, **Bonuses**, and **Expenses**.
- **Live Real-time Preview**: Shows calculated `Paid Hours`, `Gross Salary`, `Estimated Tax`, and `Estimated Net Take-Home` in real time as you enter shift times!
- **1-Click Duplicate**: Easily duplicate frequent shifts for today or any target date.

---

### 2. 🏢 Multi-Employer & Role Management with Rate Snapshots
- **Employer Configuration**: Configure Agra Tandoori, Grand Hôtel, Backup Job, or any unlimited new employers.
- **Multiple Roles**: Add separate job roles per employer (e.g. *Event Runner* @ 153.64 kr/hr, *Banquet Captain* @ 165 kr/hr).
- **Gross vs. Net Wage Types**: Explicitly tag whether hourly rates are gross or net.
- **Historical Rate Snapshots**: When you update an employer's rate in settings, **all past logged shifts preserve their original historical rate**, preventing corruption of your past income history.

---

### 3. 📊 Executive Dashboard & Analytics
- **High-Priority KPIs**:
  - 💰 **Estimated Net Pay** (Take-home income in SEK `kr`)
  - 💼 **Total Gross Salary**
  - ⏱️ **Total Paid Hours & Shift Count**
  - 🎁 **Total Tips & Supplements**
  - 📈 **Average Net Rate per Hour**
- **Date Range Filters**: Filter by `Today`, `This Week`, `This Month`, `This Year`, or `All Time`.
- **Month-over-Month Comparison**: Track your current month vs previous month earnings.
- **Interactive Charts**: Monthly income trend bar chart & employer income share.

---

### 4. 🧾 Payslip Reconciliation (Expected vs. Actual Pay)
- Log actual payslips received from your employer (*Pay Period*, *Payment Date*, *Actual Gross*, *Tax Withheld*, *Actual Net*).
- **Automatic Reconciliation**: Automatically sums all logged shifts within that pay period and calculates the exact difference (*Variance*, e.g., `-330.00 kr`), allowing you to investigate unpaid hours or tax discrepancies immediately!

---

### 5. 📅 Interactive Monthly Calendar View
- Visual monthly calendar grid showing worked shifts on each day with daily hours & earnings.
- Click on any calendar day to quick-add a shift for that date!

---

### 6. 📁 Data Persistence, Excel Export & Demo Reset
- **SQLite Database**: All records are saved securely in `work_manager.db`.
- **1-Click Excel Export**: Download your full shift log directly to a formatted Excel file (`Work_Shifts_Export.xlsx`).
- **Demo Data Reset**: Reset demo shifts anytime with 1 click.

---

## 🛠️ How to Access & Run

### Local Web Server
The server is currently running live on your system:
- 🌐 **Web App URL**: [http://127.0.0.1:5000](http://127.0.0.1:5000)
- 📊 **Dashboard API**: [http://127.0.0.1:5000/api/dashboard](http://127.0.0.1:5000/api/dashboard)
- 📁 **Excel Export**: [http://127.0.0.1:5000/api/export/excel](http://127.0.0.1:5000/api/export/excel)

### Server Execution Command:
To launch the web app backend anytime:
```bash
python C:\Users\gbrot\.gemini\antigravity\scratch\work_time_wage_tracker\app\server.py
```

import os
import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

def create_work_hours_tracker():
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    # Color Palette (Modern Executive Dark Slate & Emerald Navy)
    HEADER_FILL = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid") # Dark Slate
    HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    
    SUBHEADER_FILL = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    SUBHEADER_FONT = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

    TITLE_FONT = Font(name="Calibri", size=16, bold=True, color="0F172A")
    SECTION_FONT = Font(name="Calibri", size=13, bold=True, color="1E293B")
    
    KPI_TITLE_FONT = Font(name="Calibri", size=9, bold=True, color="64748B")
    KPI_VALUE_FONT = Font(name="Calibri", size=18, bold=True, color="0F2942")
    
    BOLD_FONT = Font(name="Calibri", size=11, bold=True, color="0F172A")
    REGULAR_FONT = Font(name="Calibri", size=11, color="334155")
    ITALIC_FONT = Font(name="Calibri", size=10, italic=True, color="64748B")

    # Fills
    ZEBRA_FILL = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    KPI_BG_FILL = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    TOTAL_FILL = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    
    AGRA_FILL = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    GRAND_FILL = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")
    BACKUP_FILL = PatternFill(start_color="F3E8FF", end_color="F3E8FF", fill_type="solid")

    # Borders
    THIN_GRAY = Side(style='thin', color='CBD5E1')
    THICK_DARK = Side(style='medium', color='1E293B')
    DOUBLE_BOTTOM = Side(style='double', color='1E293B')
    
    CARD_BORDER = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)
    CELL_BORDER = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)
    HEADER_BORDER = Border(left=THIN_GRAY, right=THIN_GRAY, top=THICK_DARK, bottom=THICK_DARK)
    TOTAL_BORDER = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=DOUBLE_BOTTOM)

    # Number Formats
    SEK_FORMAT = '#,##0.00 "kr"'
    HOURS_FORMAT = '0.00 "hrs"'
    PCT_FORMAT = '0.00%'
    DATE_FORMAT = 'YYYY-MM-DD'
    TIME_FORMAT = 'hh:mm'
    INT_FORMAT = '#,##0'

    # ---------------------------------------------------------
    # 1. JOB SETTINGS SHEET (Reference Table)
    # ---------------------------------------------------------
    ws_jobs = wb.create_sheet(title="Job_Settings")
    ws_jobs.views.sheetView[0].showGridLines = True

    ws_jobs["A1"] = "Job / Employer Settings & Tax Configuration"
    ws_jobs["A1"].font = TITLE_FONT
    ws_jobs["A2"] = "Edit hourly rates and tax percentages here. The Daily Log will automatically reference these values."
    ws_jobs["A2"].font = ITALIC_FONT

    job_headers = ["Job / Employer", "Hourly Rate (kr)", "Tax Rate (%)", "Tax Treatment Notes"]
    for col_num, h in enumerate(job_headers, 1):
        cell = ws_jobs.cell(row=4, column=col_num)
        cell.value = h
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center" if col_num > 1 else "left", vertical="center")
        cell.border = HEADER_BORDER

    job_data = [
        ("Agra Tandoori", 95.00, 0.32, "95 kr/hr net/gross with employer tax"),
        ("Grand Hotel", 153.64, 0.32, "153.64 kr/hr total salary"),
        ("Backup Job", 120.00, 0.32, "Secondary / Backup shift position"),
        ("Other / Special", 100.00, 0.30, "Custom or freelance shift")
    ]

    for row_idx, data in enumerate(job_data, 5):
        ws_jobs.cell(row=row_idx, column=1, value=data[0]).font = BOLD_FONT
        ws_jobs.cell(row=row_idx, column=1).border = CELL_BORDER
        
        c2 = ws_jobs.cell(row=row_idx, column=2, value=data[1])
        c2.font = REGULAR_FONT
        c2.number_format = SEK_FORMAT
        c2.alignment = Alignment(horizontal="right")
        c2.border = CELL_BORDER
        
        c3 = ws_jobs.cell(row=row_idx, column=3, value=data[2])
        c3.font = REGULAR_FONT
        c3.number_format = PCT_FORMAT
        c3.alignment = Alignment(horizontal="right")
        c3.border = CELL_BORDER
        
        c4 = ws_jobs.cell(row=row_idx, column=4, value=data[3])
        c4.font = ITALIC_FONT
        c4.border = CELL_BORDER

    # ---------------------------------------------------------
    # 2. DAILY LOG SHEET
    # ---------------------------------------------------------
    ws_log = wb.create_sheet(title="Daily_Log")
    ws_log.views.sheetView[0].showGridLines = True

    ws_log["A1"] = "Daily Work & Wage Log"
    ws_log["A1"].font = TITLE_FONT
    ws_log["A2"] = "Select your employer, enter shift start/end times and unpaid breaks. Rates & earnings calculate automatically!"
    ws_log["A2"].font = ITALIC_FONT

    log_headers = [
        "Date", "Job / Employer", "Start Time", "End Time", "Unpaid Break (mins)",
        "Hourly Rate (kr)", "Tax Rate (%)", "Total Paid Hours", "Gross Earnings (kr)",
        "Tax Amount (kr)", "Net Earnings (kr)", "Week #", "Year", "Month", "Shift Notes"
    ]

    for col_num, h in enumerate(log_headers, 1):
        cell = ws_log.cell(row=4, column=col_num)
        cell.value = h
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = HEADER_BORDER

    ws_log.row_dimensions[4].height = 28

    # Data validation for Job selection dropdown
    dv = DataValidation(type="list", formula1="Job_Settings!$A$5:$A$15", allow_blank=True)
    ws_log.add_data_validation(dv)
    dv.add("B5:B1000")

    # Sample shift data (realistic entries across December 2025 and January/February 2026)
    sample_shifts = [
        ("2025-12-01", "Agra Tandoori", "11:00", "19:00", 30, "Lunch shift"),
        ("2025-12-02", "Agra Tandoori", "16:00", "23:00", 30, "Dinner shift"),
        ("2025-12-04", "Grand Hotel", "07:00", "15:30", 30, "Breakfast & Lunch service"),
        ("2025-12-05", "Grand Hotel", "15:00", "23:30", 45, "Banquet evening shift"),
        ("2025-12-06", "Agra Tandoori", "12:00", "22:00", 60, "Weekend full day"),
        ("2025-12-08", "Backup Job", "09:00", "17:00", 30, "Covering shift"),
        ("2025-12-10", "Grand Hotel", "08:00", "16:30", 30, "Morning shift"),
        ("2025-12-12", "Agra Tandoori", "16:00", "23:30", 30, "Friday night service"),
        ("2025-12-15", "Agra Tandoori", "11:00", "19:00", 30, "Mid-week shift"),
        ("2025-12-18", "Grand Hotel", "14:00", "22:00", 30, "Catering event"),
        ("2025-12-20", "Agra Tandoori", "12:00", "22:30", 60, "Pre-Christmas rush"),
        ("2025-12-22", "Grand Hotel", "07:00", "15:00", 30, "Morning prep"),
        ("2025-12-27", "Agra Tandoori", "16:00", "23:00", 30, "Post-Xmas dinner"),
        
        ("2026-01-05", "Grand Hotel", "08:00", "16:30", 30, "New Year shift"),
        ("2026-01-07", "Agra Tandoori", "11:30", "19:30", 30, "Day shift"),
        ("2026-01-09", "Agra Tandoori", "16:00", "23:00", 30, "Friday night"),
        ("2026-01-12", "Grand Hotel", "07:30", "16:00", 30, "Morning service"),
        ("2026-01-15", "Backup Job", "10:00", "18:00", 30, "Event setup"),
        ("2026-01-18", "Agra Tandoori", "12:00", "21:00", 45, "Sunday service"),
        ("2026-01-22", "Grand Hotel", "15:00", "23:00", 30, "Dinner conference"),
        ("2026-01-26", "Agra Tandoori", "16:00", "23:00", 30, "Monday dinner"),

        ("2026-02-02", "Grand Hotel", "08:00", "16:30", 30, "Feb morning shift"),
        ("2026-02-04", "Agra Tandoori", "11:00", "19:00", 30, "Feb lunch shift"),
        ("2026-02-07", "Grand Hotel", "14:00", "22:30", 30, "Feb weekend event"),
    ]

    for row_idx, shift in enumerate(sample_shifts, 5):
        d_val, job_val, s_time, e_time, break_val, note_val = shift
        
        # A: Date
        c_a = ws_log.cell(row=row_idx, column=1, value=d_val)
        c_a.alignment = Alignment(horizontal="center")
        c_a.number_format = DATE_FORMAT
        
        # B: Job
        c_b = ws_log.cell(row=row_idx, column=2, value=job_val)
        c_b.alignment = Alignment(horizontal="left")
        
        # C: Start Time
        c_c = ws_log.cell(row=row_idx, column=3, value=s_time)
        c_c.alignment = Alignment(horizontal="center")
        
        # D: End Time
        c_d = ws_log.cell(row=row_idx, column=4, value=e_time)
        c_d.alignment = Alignment(horizontal="center")
        
        # E: Unpaid Break
        c_e = ws_log.cell(row=row_idx, column=5, value=break_val)
        c_e.alignment = Alignment(horizontal="right")
        c_e.number_format = INT_FORMAT

        # F: Hourly Rate Formula (VLOOKUP from Job_Settings)
        c_f = ws_log.cell(row=row_idx, column=6, value=f'=IF(B{row_idx}="","",VLOOKUP(B{row_idx},Job_Settings!$A$5:$C$15,2,FALSE))')
        c_f.alignment = Alignment(horizontal="right")
        c_f.number_format = SEK_FORMAT

        # G: Tax Rate Formula (VLOOKUP from Job_Settings)
        c_g = ws_log.cell(row=row_idx, column=7, value=f'=IF(B{row_idx}="","",VLOOKUP(B{row_idx},Job_Settings!$A$5:$C$15,3,FALSE))')
        c_g.alignment = Alignment(horizontal="right")
        c_g.number_format = PCT_FORMAT

        # H: Total Paid Hours Formula: (End - Start)*24 - Break/60
        # Handles TIME format or string format gracefully using Excel TIMEVALUE
        c_h = ws_log.cell(
            row=row_idx, column=8,
            value=f'=IF(OR(C{row_idx}="",D{row_idx}=""),"",ROUND((IF(ISNUMBER(D{row_idx}),D{row_idx},TIMEVALUE(D{row_idx}))-IF(ISNUMBER(C{row_idx}),C{row_idx},TIMEVALUE(C{row_idx})))*24-(E{row_idx}/60), 2))'
        )
        c_h.alignment = Alignment(horizontal="right")
        c_h.number_format = HOURS_FORMAT

        # I: Gross Earnings Formula: Hours * Rate
        c_i = ws_log.cell(row=row_idx, column=9, value=f'=IF(OR(H{row_idx}="",F{row_idx}=""),"",ROUND(H{row_idx}*F{row_idx}, 2))')
        c_i.alignment = Alignment(horizontal="right")
        c_i.number_format = SEK_FORMAT

        # J: Tax Amount Formula: Gross * Tax Rate
        c_j = ws_log.cell(row=row_idx, column=10, value=f'=IF(OR(I{row_idx}="",G{row_idx}=""),"",ROUND(I{row_idx}*G{row_idx}, 2))')
        c_j.alignment = Alignment(horizontal="right")
        c_j.number_format = SEK_FORMAT

        # K: Net Earnings Formula: Gross - Tax
        c_k = ws_log.cell(row=row_idx, column=11, value=f'=IF(OR(I{row_idx}="",J{row_idx}=""),"",ROUND(I{row_idx}-J{row_idx}, 2))')
        c_k.alignment = Alignment(horizontal="right")
        c_k.number_format = SEK_FORMAT

        # L: Week Number Formula
        c_l = ws_log.cell(row=row_idx, column=12, value=f'=IF(A{row_idx}="","",ISOWEEKNUM(DATEVALUE(TEXT(A{row_idx},"yyyy-mm-dd"))))')
        c_l.alignment = Alignment(horizontal="center")
        c_l.number_format = INT_FORMAT

        # M: Year Formula
        c_m = ws_log.cell(row=row_idx, column=13, value=f'=IF(A{row_idx}="","",YEAR(DATEVALUE(TEXT(A{row_idx},"yyyy-mm-dd"))))')
        c_m.alignment = Alignment(horizontal="center")
        c_m.number_format = INT_FORMAT

        # N: Month Formula
        c_n = ws_log.cell(row=row_idx, column=14, value=f'=IF(A{row_idx}="","",TEXT(DATEVALUE(TEXT(A{row_idx},"yyyy-mm-dd")),"yyyy-mm"))')
        c_n.alignment = Alignment(horizontal="center")

        # O: Notes
        c_o = ws_log.cell(row=row_idx, column=15, value=note_val)

        # Styling
        fill = ZEBRA_FILL if row_idx % 2 == 0 else WHITE_FILL
        for col_i in range(1, 16):
            cell_elem = ws_log.cell(row=row_idx, column=col_i)
            cell_elem.font = REGULAR_FONT
            cell_elem.fill = fill
            cell_elem.border = CELL_BORDER

    # Add 50 empty pre-formatted template rows for easy data entry
    start_blank = len(sample_shifts) + 5
    for row_idx in range(start_blank, start_blank + 50):
        ws_log.cell(row=row_idx, column=1).alignment = Alignment(horizontal="center")
        ws_log.cell(row=row_idx, column=1).number_format = DATE_FORMAT
        
        ws_log.cell(row=row_idx, column=5, value=30).number_format = INT_FORMAT # default 30 min break
        ws_log.cell(row=row_idx, column=5).alignment = Alignment(horizontal="right")

        ws_log.cell(row=row_idx, column=6, value=f'=IF(B{row_idx}="","",VLOOKUP(B{row_idx},Job_Settings!$A$5:$C$15,2,FALSE))').number_format = SEK_FORMAT
        ws_log.cell(row=row_idx, column=7, value=f'=IF(B{row_idx}="","",VLOOKUP(B{row_idx},Job_Settings!$A$5:$C$15,3,FALSE))').number_format = PCT_FORMAT
        ws_log.cell(row=row_idx, column=8, value=f'=IF(OR(C{row_idx}="",D{row_idx}=""),"",ROUND((IF(ISNUMBER(D{row_idx}),D{row_idx},TIMEVALUE(D{row_idx}))-IF(ISNUMBER(C{row_idx}),C{row_idx},TIMEVALUE(C{row_idx})))*24-(E{row_idx}/60), 2))').number_format = HOURS_FORMAT
        ws_log.cell(row=row_idx, column=9, value=f'=IF(OR(H{row_idx}="",F{row_idx}=""),"",ROUND(H{row_idx}*F{row_idx}, 2))').number_format = SEK_FORMAT
        ws_log.cell(row=row_idx, column=10, value=f'=IF(OR(I{row_idx}="",G{row_idx}=""),"",ROUND(I{row_idx}*G{row_idx}, 2))').number_format = SEK_FORMAT
        ws_log.cell(row=row_idx, column=11, value=f'=IF(OR(I{row_idx}="",J{row_idx}=""),"",ROUND(I{row_idx}-J{row_idx}, 2))').number_format = SEK_FORMAT
        ws_log.cell(row=row_idx, column=12, value=f'=IF(A{row_idx}="","",ISOWEEKNUM(DATEVALUE(TEXT(A{row_idx},"yyyy-mm-dd"))))').number_format = INT_FORMAT
        ws_log.cell(row=row_idx, column=13, value=f'=IF(A{row_idx}="","",YEAR(DATEVALUE(TEXT(A{row_idx},"yyyy-mm-dd"))))').number_format = INT_FORMAT
        ws_log.cell(row=row_idx, column=14, value=f'=IF(A{row_idx}="","",TEXT(DATEVALUE(TEXT(A{row_idx},"yyyy-mm-dd")),"yyyy-mm"))')

        for col_i in range(1, 16):
            cell_elem = ws_log.cell(row=row_idx, column=col_i)
            cell_elem.font = REGULAR_FONT
            cell_elem.border = CELL_BORDER

    # ---------------------------------------------------------
    # 3. WEEKLY SUMMARY SHEET
    # ---------------------------------------------------------
    ws_week = wb.create_sheet(title="Weekly_Summary")
    ws_week.views.sheetView[0].showGridLines = True

    ws_week["A1"] = "Weekly Earnings & Hours Summary"
    ws_week["A1"].font = TITLE_FONT
    ws_week["A2"] = "Automatically aggregates total shifts, paid hours, gross pay, and net take-home pay by ISO Week."
    ws_week["A2"].font = ITALIC_FONT

    week_headers = ["Year", "Week #", "Shifts Worked", "Total Paid Hours", "Gross Earnings (kr)", "Tax Amount (kr)", "Net Take-Home (kr)", "Avg Net / Hour"]
    for col_num, h in enumerate(week_headers, 1):
        cell = ws_week.cell(row=4, column=col_num)
        cell.value = h
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = HEADER_BORDER

    # Generate weeks for 2025 & 2026
    week_rows = []
    for y in [2025, 2026]:
        max_w = 52 if y == 2025 else 52
        for w in range(1, 53):
            # We only show weeks that have data or recent weeks
            if (y == 2025 and w >= 48) or (y == 2026 and w <= 12):
                week_rows.append((y, w))

    for row_idx, (y, w) in enumerate(week_rows, 5):
        # Year
        c1 = ws_week.cell(row=row_idx, column=1, value=y)
        c1.alignment = Alignment(horizontal="center")
        c1.number_format = INT_FORMAT
        
        # Week #
        c2 = ws_week.cell(row=row_idx, column=2, value=w)
        c2.alignment = Alignment(horizontal="center")
        c2.number_format = INT_FORMAT
        
        # Shifts Worked
        c3 = ws_week.cell(row=row_idx, column=3, value=f'=COUNTIFS(Daily_Log!$M$5:$M$500, A{row_idx}, Daily_Log!$L$5:$L$500, B{row_idx})')
        c3.alignment = Alignment(horizontal="center")
        c3.number_format = INT_FORMAT
        
        # Total Paid Hours
        c4 = ws_week.cell(row=row_idx, column=4, value=f'=SUMIFS(Daily_Log!$H$5:$H$500, Daily_Log!$M$5:$M$500, A{row_idx}, Daily_Log!$L$5:$L$500, B{row_idx})')
        c4.alignment = Alignment(horizontal="right")
        c4.number_format = HOURS_FORMAT
        
        # Gross Earnings
        c5 = ws_week.cell(row=row_idx, column=5, value=f'=SUMIFS(Daily_Log!$I$5:$I$500, Daily_Log!$M$5:$M$500, A{row_idx}, Daily_Log!$L$5:$L$500, B{row_idx})')
        c5.alignment = Alignment(horizontal="right")
        c5.number_format = SEK_FORMAT
        
        # Tax Amount
        c6 = ws_week.cell(row=row_idx, column=6, value=f'=SUMIFS(Daily_Log!$J$5:$J$500, Daily_Log!$M$5:$M$500, A{row_idx}, Daily_Log!$L$5:$L$500, B{row_idx})')
        c6.alignment = Alignment(horizontal="right")
        c6.number_format = SEK_FORMAT
        
        # Net Take-Home
        c7 = ws_week.cell(row=row_idx, column=7, value=f'=SUMIFS(Daily_Log!$K$5:$K$500, Daily_Log!$M$5:$M$500, A{row_idx}, Daily_Log!$L$5:$L$500, B{row_idx})')
        c7.alignment = Alignment(horizontal="right")
        c7.number_format = SEK_FORMAT

        # Avg Net / Hour
        c8 = ws_week.cell(row=row_idx, column=8, value=f'=IF(D{row_idx}>0, G{row_idx}/D{row_idx}, 0)')
        c8.alignment = Alignment(horizontal="right")
        c8.number_format = SEK_FORMAT

        fill = ZEBRA_FILL if row_idx % 2 == 0 else WHITE_FILL
        for col_i in range(1, 9):
            cell_elem = ws_week.cell(row=row_idx, column=col_i)
            cell_elem.font = REGULAR_FONT
            cell_elem.fill = fill
            cell_elem.border = CELL_BORDER

    # Total row for Weekly Summary
    tot_row = len(week_rows) + 5
    ws_week.cell(row=tot_row, column=1, value="Total").font = BOLD_FONT
    ws_week.cell(row=tot_row, column=3, value=f'=SUM(C5:C{tot_row-1})').number_format = INT_FORMAT
    ws_week.cell(row=tot_row, column=4, value=f'=SUM(D5:D{tot_row-1})').number_format = HOURS_FORMAT
    ws_week.cell(row=tot_row, column=5, value=f'=SUM(E5:E{tot_row-1})').number_format = SEK_FORMAT
    ws_week.cell(row=tot_row, column=6, value=f'=SUM(F5:F{tot_row-1})').number_format = SEK_FORMAT
    ws_week.cell(row=tot_row, column=7, value=f'=SUM(G5:G{tot_row-1})').number_format = SEK_FORMAT
    ws_week.cell(row=tot_row, column=8, value=f'=IF(D{tot_row}>0, G{tot_row}/D{tot_row}, 0)').number_format = SEK_FORMAT
    for col_i in range(1, 9):
        cell_elem = ws_week.cell(row=tot_row, column=col_i)
        cell_elem.font = BOLD_FONT
        cell_elem.fill = TOTAL_FILL
        cell_elem.border = TOTAL_BORDER

    # ---------------------------------------------------------
    # 4. MONTHLY SUMMARY SHEET
    # ---------------------------------------------------------
    ws_month = wb.create_sheet(title="Monthly_Summary")
    ws_month.views.sheetView[0].showGridLines = True

    ws_month["A1"] = "Monthly Hours & Income Breakdown"
    ws_month["A1"].font = TITLE_FONT
    ws_month["A2"] = "Tracks total hours, gross pay, tax deductions, and net income by calendar month."
    ws_month["A2"].font = ITALIC_FONT

    month_headers = ["Month Key", "Year", "Month Name", "Days Worked", "Total Paid Hours", "Gross Earnings (kr)", "Tax Deduction (kr)", "Net Take-Home (kr)", "Effective Hourly Rate"]
    for col_num, h in enumerate(month_headers, 1):
        cell = ws_month.cell(row=4, column=col_num)
        cell.value = h
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = HEADER_BORDER

    months_list = [
        ("2025-11", 2025, "November"),
        ("2025-12", 2025, "December"),
        ("2026-01", 2026, "January"),
        ("2026-02", 2026, "February"),
        ("2026-03", 2026, "March"),
        ("2026-04", 2026, "April"),
        ("2026-05", 2026, "May"),
        ("2026-06", 2026, "June"),
        ("2026-07", 2026, "July"),
        ("2026-08", 2026, "August"),
        ("2026-09", 2026, "September"),
        ("2026-10", 2026, "October"),
        ("2026-11", 2026, "November"),
        ("2026-12", 2026, "December")
    ]

    for row_idx, (m_key, yr, m_name) in enumerate(months_list, 5):
        # Month Key
        c1 = ws_month.cell(row=row_idx, column=1, value=m_key)
        c1.alignment = Alignment(horizontal="center")
        
        # Year
        c2 = ws_month.cell(row=row_idx, column=2, value=yr)
        c2.alignment = Alignment(horizontal="center")
        c2.number_format = INT_FORMAT
        
        # Month Name
        c3 = ws_month.cell(row=row_idx, column=3, value=m_name)
        c3.alignment = Alignment(horizontal="left")
        
        # Days Worked
        c4 = ws_month.cell(row=row_idx, column=4, value=f'=COUNTIFS(Daily_Log!$N$5:$N$500, A{row_idx})')
        c4.alignment = Alignment(horizontal="center")
        c4.number_format = INT_FORMAT

        # Total Paid Hours
        c5 = ws_month.cell(row=row_idx, column=5, value=f'=SUMIFS(Daily_Log!$H$5:$H$500, Daily_Log!$N$5:$N$500, A{row_idx})')
        c5.alignment = Alignment(horizontal="right")
        c5.number_format = HOURS_FORMAT

        # Gross Earnings
        c6 = ws_month.cell(row=row_idx, column=6, value=f'=SUMIFS(Daily_Log!$I$5:$I$500, Daily_Log!$N$5:$N$500, A{row_idx})')
        c6.alignment = Alignment(horizontal="right")
        c6.number_format = SEK_FORMAT

        # Tax Deduction
        c7 = ws_month.cell(row=row_idx, column=7, value=f'=SUMIFS(Daily_Log!$J$5:$J$500, Daily_Log!$N$5:$N$500, A{row_idx})')
        c7.alignment = Alignment(horizontal="right")
        c7.number_format = SEK_FORMAT

        # Net Take-Home
        c8 = ws_month.cell(row=row_idx, column=8, value=f'=SUMIFS(Daily_Log!$K$5:$K$500, Daily_Log!$N$5:$N$500, A{row_idx})')
        c8.alignment = Alignment(horizontal="right")
        c8.number_format = SEK_FORMAT

        # Effective Hourly Rate
        c9 = ws_month.cell(row=row_idx, column=9, value=f'=IF(E{row_idx}>0, H{row_idx}/E{row_idx}, 0)')
        c9.alignment = Alignment(horizontal="right")
        c9.number_format = SEK_FORMAT

        fill = ZEBRA_FILL if row_idx % 2 == 0 else WHITE_FILL
        for col_i in range(1, 10):
            cell_elem = ws_month.cell(row=row_idx, column=col_i)
            cell_elem.font = REGULAR_FONT
            cell_elem.fill = fill
            cell_elem.border = CELL_BORDER

    tot_m_row = len(months_list) + 5
    ws_month.cell(row=tot_m_row, column=1, value="Total").font = BOLD_FONT
    ws_month.cell(row=tot_m_row, column=4, value=f'=SUM(D5:D{tot_m_row-1})').number_format = INT_FORMAT
    ws_month.cell(row=tot_m_row, column=5, value=f'=SUM(E5:E{tot_m_row-1})').number_format = HOURS_FORMAT
    ws_month.cell(row=tot_m_row, column=6, value=f'=SUM(F5:F{tot_m_row-1})').number_format = SEK_FORMAT
    ws_month.cell(row=tot_m_row, column=7, value=f'=SUM(G5:G{tot_m_row-1})').number_format = SEK_FORMAT
    ws_month.cell(row=tot_m_row, column=8, value=f'=SUM(H5:H{tot_m_row-1})').number_format = SEK_FORMAT
    ws_month.cell(row=tot_m_row, column=9, value=f'=IF(E{tot_m_row}>0, H{tot_m_row}/E{tot_m_row}, 0)').number_format = SEK_FORMAT
    for col_i in range(1, 10):
        cell_elem = ws_month.cell(row=tot_m_row, column=col_i)
        cell_elem.font = BOLD_FONT
        cell_elem.fill = TOTAL_FILL
        cell_elem.border = TOTAL_BORDER

    # ---------------------------------------------------------
    # 5. YEARLY SUMMARY SHEET
    # ---------------------------------------------------------
    ws_year = wb.create_sheet(title="Yearly_Summary")
    ws_year.views.sheetView[0].showGridLines = True

    ws_year["A1"] = "Yearly Income & Hours Overview"
    ws_year["A1"].font = TITLE_FONT
    ws_year["A2"] = "High-level annual performance totals, days worked, and average daily/hourly rates."
    ws_year["A2"].font = ITALIC_FONT

    year_headers = ["Year", "Total Days Worked", "Total Paid Hours", "Gross Earnings (kr)", "Tax Amount (kr)", "Net Take-Home (kr)", "Avg Daily Net (kr)", "Avg Hourly Net (kr)"]
    for col_num, h in enumerate(year_headers, 1):
        cell = ws_year.cell(row=4, column=col_num)
        cell.value = h
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = HEADER_BORDER

    years_list = [2025, 2026, 2027]
    for row_idx, yr in enumerate(years_list, 5):
        # Year
        c1 = ws_year.cell(row=row_idx, column=1, value=yr)
        c1.alignment = Alignment(horizontal="center")
        c1.number_format = INT_FORMAT

        # Days Worked
        c2 = ws_year.cell(row=row_idx, column=2, value=f'=COUNTIFS(Daily_Log!$M$5:$M$500, A{row_idx})')
        c2.alignment = Alignment(horizontal="center")
        c2.number_format = INT_FORMAT

        # Total Paid Hours
        c3 = ws_year.cell(row=row_idx, column=3, value=f'=SUMIFS(Daily_Log!$H$5:$H$500, Daily_Log!$M$5:$M$500, A{row_idx})')
        c3.alignment = Alignment(horizontal="right")
        c3.number_format = HOURS_FORMAT

        # Gross Earnings
        c4 = ws_year.cell(row=row_idx, column=4, value=f'=SUMIFS(Daily_Log!$I$5:$I$500, Daily_Log!$M$5:$M$500, A{row_idx})')
        c4.alignment = Alignment(horizontal="right")
        c4.number_format = SEK_FORMAT

        # Tax Amount
        c5 = ws_year.cell(row=row_idx, column=5, value=f'=SUMIFS(Daily_Log!$J$5:$J$500, Daily_Log!$M$5:$M$500, A{row_idx})')
        c5.alignment = Alignment(horizontal="right")
        c5.number_format = SEK_FORMAT

        # Net Take-Home
        c6 = ws_year.cell(row=row_idx, column=6, value=f'=SUMIFS(Daily_Log!$K$5:$K$500, Daily_Log!$M$5:$M$500, A{row_idx})')
        c6.alignment = Alignment(horizontal="right")
        c6.number_format = SEK_FORMAT

        # Avg Daily Net
        c7 = ws_year.cell(row=row_idx, column=7, value=f'=IF(B{row_idx}>0, F{row_idx}/B{row_idx}, 0)')
        c7.alignment = Alignment(horizontal="right")
        c7.number_format = SEK_FORMAT

        # Avg Hourly Net
        c8 = ws_year.cell(row=row_idx, column=8, value=f'=IF(C{row_idx}>0, F{row_idx}/C{row_idx}, 0)')
        c8.alignment = Alignment(horizontal="right")
        c8.number_format = SEK_FORMAT

        fill = ZEBRA_FILL if row_idx % 2 == 0 else WHITE_FILL
        for col_i in range(1, 9):
            cell_elem = ws_year.cell(row=row_idx, column=col_i)
            cell_elem.font = REGULAR_FONT
            cell_elem.fill = fill
            cell_elem.border = CELL_BORDER

    # ---------------------------------------------------------
    # 6. EXECUTIVE DASHBOARD SHEET
    # ---------------------------------------------------------
    ws_dash = wb.create_sheet(title="Dashboard")
    ws_dash.views.sheetView[0].showGridLines = True

    # Title block
    ws_dash["A1"] = "💼 WORK HOURS & WAGE DASHBOARD"
    ws_dash["A1"].font = Font(name="Calibri", size=18, bold=True, color="0F172A")
    ws_dash["A2"] = "Real-Time Key Performance Indicators & Multi-Job Earnings Summary (Currency: SEK kr)"
    ws_dash["A2"].font = ITALIC_FONT

    # KPI Cards Layout (Rows 4 to 6)
    # Card 1: Total Net Earnings (Cols B:C)
    # Card 2: Total Gross Earnings (Cols D:E)
    # Card 3: Total Paid Hours (Cols F:G)
    # Card 4: Total Days Worked (Cols H:I)
    
    kpis = [
        ("TOTAL NET TAKE-HOME", "=SUM(Daily_Log!K5:K500)", SEK_FORMAT, "B", "C"),
        ("TOTAL GROSS EARNINGS", "=SUM(Daily_Log!I5:I500)", SEK_FORMAT, "E", "F"),
        ("TOTAL HOURS WORKED", "=SUM(Daily_Log!H5:H500)", HOURS_FORMAT, "H", "I"),
        ("TOTAL SHIFTS WORKED", "=COUNTA(Daily_Log!A5:A500)", INT_FORMAT, "K", "L")
    ]

    for title, formula, fmt, col1, col2 in kpis:
        c_title = ws_dash[f"{col1}4"]
        c_title.value = title
        c_title.font = KPI_TITLE_FONT
        c_title.alignment = Alignment(horizontal="center", vertical="center")

        c_val = ws_dash[f"{col1}5"]
        c_val.value = formula
        c_val.font = KPI_VALUE_FONT
        c_val.number_format = fmt
        c_val.alignment = Alignment(horizontal="center", vertical="center")

        # Merge 2 columns for card width
        ws_dash.merge_cells(f"{col1}4:{col2}4")
        ws_dash.merge_cells(f"{col1}5:{col2}6")

        for r in range(4, 7):
            for c_letter in [col1, col2]:
                cell = ws_dash[f"{c_letter}{r}"]
                cell.fill = KPI_BG_FILL
                cell.border = CARD_BORDER

    # Section 1: Employer / Job Breakdown Table (Row 8)
    ws_dash["B8"] = "📊 Multi-Job Earnings & Hours Breakdown"
    ws_dash["B8"].font = SECTION_FONT

    job_dash_headers = ["Job / Employer", "Configured Rate", "Tax Rate", "Total Hours Worked", "Gross Earnings (kr)", "Net Take-Home (kr)", "Share of Net Income"]
    for idx, h in enumerate(job_dash_headers, 2): # Start column B
        cell = ws_dash.cell(row=10, column=idx)
        cell.value = h
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = HEADER_BORDER

    for r_idx, job in enumerate(["Agra Tandoori", "Grand Hotel", "Backup Job", "Other / Special"], 11):
        # Job Name
        ws_dash.cell(row=r_idx, column=2, value=job).font = BOLD_FONT
        ws_dash.cell(row=r_idx, column=2).border = CELL_BORDER

        # Hourly Rate
        c_r = ws_dash.cell(row=r_idx, column=3, value=f'=VLOOKUP(B{r_idx},Job_Settings!$A$5:$C$15,2,FALSE)')
        c_r.font = REGULAR_FONT
        c_r.number_format = SEK_FORMAT
        c_r.alignment = Alignment(horizontal="right")
        c_r.border = CELL_BORDER

        # Tax Rate
        c_t = ws_dash.cell(row=r_idx, column=4, value=f'=VLOOKUP(B{r_idx},Job_Settings!$A$5:$C$15,3,FALSE)')
        c_t.font = REGULAR_FONT
        c_t.number_format = PCT_FORMAT
        c_t.alignment = Alignment(horizontal="right")
        c_t.border = CELL_BORDER

        # Total Hours Worked
        c_h = ws_dash.cell(row=r_idx, column=5, value=f'=SUMIFS(Daily_Log!$H$5:$H$500, Daily_Log!$B$5:$B$500, B{r_idx})')
        c_h.font = REGULAR_FONT
        c_h.number_format = HOURS_FORMAT
        c_h.alignment = Alignment(horizontal="right")
        c_h.border = CELL_BORDER

        # Gross Earnings
        c_g = ws_dash.cell(row=r_idx, column=6, value=f'=SUMIFS(Daily_Log!$I$5:$I$500, Daily_Log!$B$5:$B$500, B{r_idx})')
        c_g.font = REGULAR_FONT
        c_g.number_format = SEK_FORMAT
        c_g.alignment = Alignment(horizontal="right")
        c_g.border = CELL_BORDER

        # Net Take-Home
        c_n = ws_dash.cell(row=r_idx, column=7, value=f'=SUMIFS(Daily_Log!$K$5:$K$500, Daily_Log!$B$5:$B$500, B{r_idx})')
        c_n.font = REGULAR_FONT
        c_n.number_format = SEK_FORMAT
        c_n.alignment = Alignment(horizontal="right")
        c_n.border = CELL_BORDER

        # Share of Net Income
        c_s = ws_dash.cell(row=r_idx, column=8, value=f'=IF($G$15>0, G{r_idx}/$G$15, 0)')
        c_s.font = REGULAR_FONT
        c_s.number_format = PCT_FORMAT
        c_s.alignment = Alignment(horizontal="right")
        c_s.border = CELL_BORDER

    # Total row for Job Breakdown
    ws_dash.cell(row=15, column=2, value="Total All Jobs").font = BOLD_FONT
    ws_dash.cell(row=15, column=5, value='=SUM(E11:E14)').number_format = HOURS_FORMAT
    ws_dash.cell(row=15, column=6, value='=SUM(F11:F14)').number_format = SEK_FORMAT
    ws_dash.cell(row=15, column=7, value='=SUM(G11:G14)').number_format = SEK_FORMAT
    ws_dash.cell(row=15, column=8, value='=SUM(H11:H14)').number_format = PCT_FORMAT
    
    for c_i in range(2, 9):
        cell_elem = ws_dash.cell(row=15, column=c_i)
        cell_elem.font = BOLD_FONT
        cell_elem.fill = TOTAL_FILL
        cell_elem.border = TOTAL_BORDER

    # Section 2: Recent Shift Log Preview (Row 18)
    ws_dash["B18"] = "📅 Recent Shifts Quick View"
    ws_dash["B18"].font = SECTION_FONT

    recent_headers = ["Date", "Job / Employer", "Hours", "Gross (kr)", "Net (kr)", "Notes"]
    for idx, h in enumerate(recent_headers, 2):
        cell = ws_dash.cell(row=20, column=idx)
        cell.value = h
        cell.font = SUBHEADER_FONT
        cell.fill = SUBHEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = HEADER_BORDER

    for r_idx, log_row in enumerate(range(5, 15), 21):
        ws_dash.cell(row=r_idx, column=2, value=f'=Daily_Log!A{log_row}').number_format = DATE_FORMAT
        ws_dash.cell(row=r_idx, column=2).alignment = Alignment(horizontal="center")
        
        ws_dash.cell(row=r_idx, column=3, value=f'=Daily_Log!B{log_row}')
        
        ws_dash.cell(row=r_idx, column=4, value=f'=Daily_Log!H{log_row}').number_format = HOURS_FORMAT
        ws_dash.cell(row=r_idx, column=4).alignment = Alignment(horizontal="right")
        
        ws_dash.cell(row=r_idx, column=5, value=f'=Daily_Log!I{log_row}').number_format = SEK_FORMAT
        ws_dash.cell(row=r_idx, column=5).alignment = Alignment(horizontal="right")
        
        ws_dash.cell(row=r_idx, column=6, value=f'=Daily_Log!K{log_row}').number_format = SEK_FORMAT
        ws_dash.cell(row=r_idx, column=6).alignment = Alignment(horizontal="right")
        
        ws_dash.cell(row=r_idx, column=7, value=f'=Daily_Log!O{log_row}')

        fill = ZEBRA_FILL if r_idx % 2 == 0 else WHITE_FILL
        for c_i in range(2, 8):
            cell_elem = ws_dash.cell(row=r_idx, column=c_i)
            cell_elem.font = REGULAR_FONT
            cell_elem.fill = fill
            cell_elem.border = CELL_BORDER

    # Section 3: Apple Calendar Sync Quick Instructions
    ws_dash["I8"] = "📱 Apple Calendar Sync Guide"
    ws_dash["I8"].font = SECTION_FONT

    instructions = [
        ("Step 1", "Create events on Apple Calendar with titles like 'Agra Tandoori' or 'Grand Hotel'."),
        ("Step 2", "Export calendar (.ics file) or copy Webcal URL from iCal."),
        ("Step 3", "Run the included `sync_apple_calendar.py` script to import shifts automatically."),
        ("Step 4", "Alternatively, use Apple Shortcuts app to send shift logs directly to Excel!")
    ]

    for idx, (step, desc) in enumerate(instructions, 10):
        c_step = ws_dash.cell(row=idx, column=9, value=step)
        c_step.font = BOLD_FONT
        c_step.fill = AGRA_FILL
        c_step.border = CELL_BORDER
        c_step.alignment = Alignment(horizontal="center")

        c_desc = ws_dash.cell(row=idx, column=10, value=desc)
        c_desc.font = REGULAR_FONT
        c_desc.border = CELL_BORDER
        ws_dash.merge_cells(start_row=idx, start_column=10, end_row=idx, end_column=12)

    # Reorder sheets so Dashboard is first tab
    sheet_order = ["Dashboard", "Daily_Log", "Weekly_Summary", "Monthly_Summary", "Yearly_Summary", "Job_Settings"]
    wb._sheets = [wb[s] for s in sheet_order]

    # Auto-adjust column widths across all sheets
    for sheet in wb.worksheets:
        for col in sheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                # Avoid long text in merged title cells expanding column A unreasonably
                if cell.row in [1, 2] and col_letter in ['A', 'B', 'I']:
                    continue
                if cell.value:
                    val_str = str(cell.value)
                    if not val_str.startswith('='):
                        max_len = max(max_len, len(val_str))
            sheet.column_dimensions[col_letter].width = max(max_len + 4, 14)

    # Special adjustments for specific columns
    ws_dash.column_dimensions['A'].width = 3
    ws_dash.column_dimensions['B'].width = 20
    ws_dash.column_dimensions['C'].width = 16
    ws_dash.column_dimensions['D'].width = 16
    ws_dash.column_dimensions['E'].width = 20
    ws_dash.column_dimensions['F'].width = 20
    ws_dash.column_dimensions['G'].width = 20
    ws_dash.column_dimensions['H'].width = 18
    ws_dash.column_dimensions['I'].width = 14
    ws_dash.column_dimensions['J'].width = 35

    ws_log.column_dimensions['A'].width = 14
    ws_log.column_dimensions['B'].width = 18
    ws_log.column_dimensions['E'].width = 18
    ws_log.column_dimensions['H'].width = 18
    ws_log.column_dimensions['I'].width = 18
    ws_log.column_dimensions['J'].width = 18
    ws_log.column_dimensions['K'].width = 18
    ws_log.column_dimensions['O'].width = 30

    output_path = r"C:\Users\gbrot\.gemini\antigravity\scratch\work_time_wage_tracker\Work_Hours_and_Wage_Tracker.xlsx"
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
    print(f"Successfully generated Work_Hours_and_Wage_Tracker.xlsx at: {output_path}")

if __name__ == "__main__":
    create_work_hours_tracker()

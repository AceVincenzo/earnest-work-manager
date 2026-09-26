import os
import re
import datetime
import openpyxl
from icalendar import Calendar

def parse_break_from_text(text):
    if not text:
        return 30 # default 30 mins
    # Search for patterns like "break: 45", "45 min break", "break 30m"
    match = re.search(r'(?:break|rest)[:\s]*(\d+)\s*(?:m|min|mins|minutes)?', text, re.IGNORECASE)
    if match:
        return int(match.group(1))
    match_mins = re.search(r'(\d+)\s*(?:m|min|mins|minutes)\s*break', text, re.IGNORECASE)
    if match_mins:
        return int(match_mins.group(1))
    return 30

def identify_job(summary):
    if not summary:
        return "Backup Job"
    s_lower = summary.lower()
    if "agra" in s_lower or "tandoori" in s_lower:
        return "Agra Tandoori"
    elif "grand" in s_lower or "hotel" in s_lower:
        return "Grand Hotel"
    elif "backup" in s_lower:
        return "Backup Job"
    else:
        # Check general work keywords
        return "Other / Special"

def sync_ics_to_excel(ics_filepath, excel_filepath):
    if not os.path.exists(ics_filepath):
        print(f"Error: Calendar file not found at '{ics_filepath}'")
        return False

    if not os.path.exists(excel_filepath):
        print(f"Error: Excel file not found at '{excel_filepath}'")
        return False

    # Load Calendar
    with open(ics_filepath, 'rb') as f:
        cal = Calendar.from_ical(f.read())

    # Load Workbook
    wb = openpyxl.load_workbook(excel_filepath)
    if "Daily_Log" not in wb.sheetnames:
        print("Error: Sheet 'Daily_Log' not found in workbook.")
        return False
    
    ws = wb["Daily_Log"]

    # Read existing entries to prevent duplicate import
    existing_entries = set()
    for row in range(5, ws.max_row + 1):
        d_val = ws.cell(row=row, column=1).value
        j_val = ws.cell(row=row, column=2).value
        s_val = ws.cell(row=row, column=3).value
        if d_val and j_val:
            # Format date as YYYY-MM-DD string
            if isinstance(d_val, (datetime.date, datetime.datetime)):
                d_str = d_val.strftime("%Y-%m-%d")
            else:
                d_str = str(d_val)[:10]
            s_str = str(s_val).strip() if s_val else ""
            existing_entries.add((d_str, j_val, s_str))

    # Find next empty row in Daily_Log
    target_row = 5
    while ws.cell(row=target_row, column=1).value is not None:
        target_row += 1

    imported_count = 0
    skipped_count = 0

    print("\n--- Processing Apple Calendar Events ---")
    for component in cal.walk():
        if component.name == "VEVENT":
            summary = str(component.get('summary', 'Work Shift'))
            description = str(component.get('description', ''))
            
            dtstart = component.get('dtstart')
            dtend = component.get('dtend')

            if not dtstart or not dtend:
                continue

            start_dt = dtstart.dt
            end_dt = dtend.dt

            # Ensure datetime format
            if isinstance(start_dt, datetime.date) and not isinstance(start_dt, datetime.datetime):
                # All-day event without time, skip or default 8h
                continue

            date_str = start_dt.strftime("%Y-%m-%d")
            start_time_str = start_dt.strftime("%H:%M")
            end_time_str = end_dt.strftime("%H:%M")

            job_name = identify_job(summary)
            break_mins = parse_break_from_text(description)
            shift_note = f"Imported from Apple Calendar: {summary}"

            # Check duplication
            if (date_str, job_name, start_time_str) in existing_entries:
                print(f"Skipped duplicate: {date_str} | {job_name} ({start_time_str}-{end_time_str})")
                skipped_count += 1
                continue

            # Append to Daily_Log
            ws.cell(row=target_row, column=1, value=date_str).alignment = openpyxl.styles.Alignment(horizontal="center")
            ws.cell(row=target_row, column=2, value=job_name).alignment = openpyxl.styles.Alignment(horizontal="left")
            ws.cell(row=target_row, column=3, value=start_time_str).alignment = openpyxl.styles.Alignment(horizontal="center")
            ws.cell(row=target_row, column=4, value=end_time_str).alignment = openpyxl.styles.Alignment(horizontal="center")
            ws.cell(row=target_row, column=5, value=break_mins).alignment = openpyxl.styles.Alignment(horizontal="right")

            # Formulas
            ws.cell(row=target_row, column=6, value=f'=IF(B{target_row}="","",VLOOKUP(B{target_row},Job_Settings!$A$5:$C$15,2,FALSE))')
            ws.cell(row=target_row, column=7, value=f'=IF(B{target_row}="","",VLOOKUP(B{target_row},Job_Settings!$A$5:$C$15,3,FALSE))')
            ws.cell(row=target_row, column=8, value=f'=IF(OR(C{target_row}="",D{target_row}=""),"",ROUND((IF(ISNUMBER(D{target_row}),D{target_row},TIMEVALUE(D{target_row}))-IF(ISNUMBER(C{target_row}),C{target_row},TIMEVALUE(C{target_row})))*24-(E{target_row}/60), 2))')
            ws.cell(row=target_row, column=9, value=f'=IF(OR(H{target_row}="",F{target_row}=""),"",ROUND(H{target_row}*F{target_row}, 2))')
            ws.cell(row=target_row, column=10, value=f'=IF(OR(I{target_row}="",G{target_row}=""),"",ROUND(I{target_row}*G{target_row}, 2))')
            ws.cell(row=target_row, column=11, value=f'=IF(OR(I{target_row}="",J{target_row}=""),"",ROUND(I{target_row}-J{target_row}, 2))')
            ws.cell(row=target_row, column=12, value=f'=IF(A{target_row}="","",ISOWEEKNUM(DATEVALUE(TEXT(A{target_row},"yyyy-mm-dd"))))')
            ws.cell(row=target_row, column=13, value=f'=IF(A{target_row}="","",YEAR(DATEVALUE(TEXT(A{target_row},"yyyy-mm-dd"))))')
            ws.cell(row=target_row, column=14, value=f'=IF(A{target_row}="","",TEXT(DATEVALUE(TEXT(A{target_row},"yyyy-mm-dd")),"yyyy-mm"))')
            ws.cell(row=target_row, column=15, value=shift_note)

            # Apply cell formatting & borders
            SEK_FORMAT = '#,##0.00 "kr"'
            HOURS_FORMAT = '0.00 "hrs"'
            PCT_FORMAT = '0.00%'
            INT_FORMAT = '#,##0'
            THIN_GRAY = openpyxl.styles.Side(style='thin', color='CBD5E1')
            BORDER = openpyxl.styles.Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)
            FONT = openpyxl.styles.Font(name="Calibri", size=11, color="334155")

            ws.cell(row=target_row, column=5).number_format = INT_FORMAT
            ws.cell(row=target_row, column=6).number_format = SEK_FORMAT
            ws.cell(row=target_row, column=7).number_format = PCT_FORMAT
            ws.cell(row=target_row, column=8).number_format = HOURS_FORMAT
            ws.cell(row=target_row, column=9).number_format = SEK_FORMAT
            ws.cell(row=target_row, column=10).number_format = SEK_FORMAT
            ws.cell(row=target_row, column=11).number_format = SEK_FORMAT

            for c_idx in range(1, 16):
                c_cell = ws.cell(row=target_row, column=c_idx)
                c_cell.font = FONT
                c_cell.border = BORDER

            print(f"[OK] Imported: {date_str} | {job_name} | {start_time_str}-{end_time_str} ({break_mins}m break)")
            existing_entries.add((date_str, job_name, start_time_str))
            imported_count += 1
            target_row += 1

    wb.save(excel_filepath)
    print(f"\n[DONE] Sync completed successfully!")
    print(f"   - Total New Shifts Imported: {imported_count}")
    print(f"   - Skipped / Duplicate Shifts: {skipped_count}")
    print(f"   - Workbook updated: {excel_filepath}")
    return True

if __name__ == "__main__":
    default_ics = r"C:\Users\gbrot\.gemini\antigravity\scratch\work_time_wage_tracker\sample_apple_calendar.ics"
    default_excel = r"C:\Users\gbrot\.gemini\antigravity\scratch\work_time_wage_tracker\Work_Hours_and_Wage_Tracker.xlsx"
    
    import sys
    ics_path = sys.argv[1] if len(sys.argv) > 1 else default_ics
    excel_path = sys.argv[2] if len(sys.argv) > 2 else default_excel

    sync_ics_to_excel(ics_path, excel_path)

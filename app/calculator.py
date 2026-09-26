import datetime

def calculate_shift(start_date_str, start_time_str, end_time_str, break_mins=30,
                    hourly_rate=95.0, rate_type='gross', tax_rate=0.32,
                    ob_pay=0.0, overtime_pay=0.0, tips=0.0, bonus=0.0, expenses=0.0):
    """
    Central deterministic calculation engine for shifts.
    Handles overnight shifts (e.g., 18:00 to 02:00), breaks, rates, supplements, and tax.
    """
    if not start_date_str or not start_time_str or not end_time_str:
        raise ValueError("Missing start date, start time, or end time.")

    # Parse start date & time
    s_date = datetime.datetime.strptime(start_date_str, "%Y-%m-%d").date()
    s_time = datetime.datetime.strptime(start_time_str, "%H:%M").time()
    e_time = datetime.datetime.strptime(end_time_str, "%H:%M").time()

    start_dt = datetime.datetime.combine(s_date, s_time)

    # If end time is before or equal to start time, it's an overnight shift!
    if e_time <= s_time:
        end_dt = datetime.datetime.combine(s_date + datetime.timedelta(days=1), e_time)
    else:
        end_dt = datetime.datetime.combine(s_date, e_time)

    # Total duration in minutes
    elapsed_seconds = (end_dt - start_dt).total_seconds()
    elapsed_minutes = elapsed_seconds / 60.0

    if break_mins < 0:
        raise ValueError("Break minutes cannot be negative.")

    # Calculate net paid minutes
    paid_minutes = max(0.0, elapsed_minutes - float(break_mins))
    paid_hours = round(paid_minutes / 60.0, 2)

    # Base pay calculation
    if rate_type == 'net':
        # If rate is defined as net rate, back-calculate gross assuming flat tax rate
        net_base = paid_hours * float(hourly_rate)
        if tax_rate < 1.0:
            gross_base = net_base / (1.0 - float(tax_rate))
        else:
            gross_base = net_base
    else:
        # Gross rate
        gross_base = paid_hours * float(hourly_rate)

    # Total gross includes supplements & tips
    total_gross = gross_base + float(ob_pay) + float(overtime_pay) + float(tips) + float(bonus) - float(expenses)
    total_gross = round(max(0.0, total_gross), 2)

    # Estimated tax & net pay
    estimated_tax = round(total_gross * float(tax_rate), 2)
    estimated_net = round(total_gross - estimated_tax, 2)

    return {
        "start_datetime": start_dt.strftime("%Y-%m-%d %H:%M"),
        "end_datetime": end_dt.strftime("%Y-%m-%d %H:%M"),
        "is_overnight": e_time <= s_time,
        "elapsed_hours": round(elapsed_minutes / 60.0, 2),
        "paid_hours": paid_hours,
        "gross_base": round(gross_base, 2),
        "total_gross": total_gross,
        "estimated_tax": estimated_tax,
        "estimated_net": estimated_net,
        "hourly_rate": float(hourly_rate),
        "tax_rate": float(tax_rate)
    }

def run_tests():
    print("Running calculation engine test suite...")
    
    # Test 1: Standard Day Shift (09:00 -> 17:00, 30m break) = 7.5h
    res1 = calculate_shift("2026-09-25", "09:00", "17:00", break_mins=30, hourly_rate=100, tax_rate=0.30)
    assert res1["paid_hours"] == 7.5, f"Expected 7.5, got {res1['paid_hours']}"
    assert res1["total_gross"] == 750.0, f"Expected 750.0, got {res1['total_gross']}"
    assert res1["estimated_net"] == 525.0, f"Expected 525.0, got {res1['estimated_net']}"
    print("  [OK] Test 1 Passed: Standard day shift (7.5h)")

    # Test 2: Overnight Shift (18:00 -> 02:00 next day, 30m break) = 7.5h
    res2 = calculate_shift("2026-09-25", "18:00", "02:00", break_mins=30, hourly_rate=153.64, tax_rate=0.32)
    assert res2["is_overnight"] == True, "Should be marked overnight"
    assert res2["paid_hours"] == 7.5, f"Expected 7.5, got {res2['paid_hours']}"
    assert res2["end_datetime"] == "2026-09-26 02:00", f"Unexpected end datetime: {res2['end_datetime']}"
    print("  [OK] Test 2 Passed: Overnight shift (18:00 -> 02:00 next day)")

    # Test 3: Shift with OB, Tips & Overtime
    res3 = calculate_shift("2026-09-25", "16:00", "23:00", break_mins=30, hourly_rate=95.0, tax_rate=0.32,
                           ob_pay=150.0, tips=200.0)
    assert res3["paid_hours"] == 6.5
    expected_gross = round(6.5 * 95.0 + 150.0 + 200.0, 2)
    assert res3["total_gross"] == expected_gross, f"Expected {expected_gross}, got {res3['total_gross']}"
    print("  [OK] Test 3 Passed: Shift with OB & Tips")

    # Test 4: Zero break
    res4 = calculate_shift("2026-09-25", "12:00", "16:00", break_mins=0, hourly_rate=100)
    assert res4["paid_hours"] == 4.0
    print("  [OK] Test 4 Passed: Zero break")

    print("All calculation tests passed successfully!\n")

if __name__ == "__main__":
    run_tests()

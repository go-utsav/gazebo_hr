from pathlib import Path

from weekly.payroll_service import build_excel_bytes, calculate_payroll, parse_employee_hours


emp = Path(r"C:\Users\utsav\projects\gazebo_clockrite_holiday_report_dot_net\prompts\test_data\31.03.2026\dgross_paysummary2.xls")
con = Path(r"C:\Users\utsav\projects\gazebo_clockrite_holiday_report_dot_net\prompts\test_data\data\employee_contract_hours.xls")

with emp.open("rb") as employee_file, con.open("rb") as contracted_file:
    rows = parse_employee_hours(employee_file)
    result = calculate_payroll(rows, contracted_file)
    output = build_excel_bytes(result)

out_path = Path(r"C:\Users\utsav\projects\gazebo\python\weekly_sample_output.xlsx")
out_path.write_bytes(output)

print("rows:", len(result.rows))
print("agency:", len(result.agency_rows))
print("gazebo:", len(result.gazebo_rows))
print("total_paid_hours:", result.total_paid_hours)
print("output:", out_path)

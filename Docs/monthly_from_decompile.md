## Monthly EXE logic (decompiled)

Source: `C:\Users\utsav\projects\gazebo_clockrite_holiday_report_dot_net\Monthly_decompiled\ClockRiteSummary\Form1.cs`

- App collects 4 weekly files (`file1..file4`) and optional week 5 (`file5`).
- Login password is hard-coded (`ENTRYCL`) in `Login`.
- Each input workbook is read from first worksheet:
  - Start date: `cell(1,5)`, end date: `cell(1,8)`.
  - Employee rows start at row 4 (`Name, Category, SageNo, Basic, MonFri, SatSun, Annual, TotalPaid`) until blank name.
  - Then adjustments block (`Name, Type, Value`) from column 2.
  - Then category totals block (`Category` in column 2, numeric bands in columns 4..8).
  - A non-agency total is read from bottom total row (`col 8`).
- For each week, data is stored in `Summary` object:
  - `employees`, `employeeTotals`, `adjustments`, `startDate`, `endDate`, `total`.

### Weekly sheet writeback

- Output workbook has `Week1..WeekN` sheets plus one `Summary` sheet.
- Each week sheet writes:
  - Week header and dates.
  - Employee table.
  - Adjustments table.
  - Category totals table.
  - Grouped rollup where keys are:
    - First 4 chars of category, or
    - For agency (`A-*`): first 4 chars + space + chars 6..9.
  - Bottom non-agency total row (excludes categories starting `A-`).

### Summary sheet aggregation rules

- Merge employees across weeks by **Name** key (`Employee.empDict[name]`):
  - Sums `BasicHours, MonFriOvertime, SatSunOvertime, AnnualHoliday, TotalPaidHours`.
  - Overwrites `SageNo` and `Category` with latest seen entry.
- Merge category totals by exact `Category`.
- `HourForm` allows editing `IsHourly` for each employee.
  - Validation: agency category (`A-`) cannot be non-hourly.
- While writing summary category totals:
  - Normal category row is written.
  - If category has non-hourly employees, an extra negative row is written:
    - `"<Category> non-hourly hours"` and negative sums in cols 4..8.
  - Final non-agency total row excludes categories starting `A-`.

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from io import BytesIO
from typing import Any

from openpyxl import Workbook


@dataclass
class MonthlyEmployee:
    Name: str
    Category: str
    SageNo: int
    BasicHours: float
    MonFriOvertime: float
    SatSunOvertime: float
    AnnualHoliday: float
    TotalPaidHours: float
    IsHourly: bool = True


@dataclass
class MonthlyEmployeeTotal:
    Category: str
    BasicHours: float
    MonFriOvertime: float
    SatSunOvertime: float
    AnnualHoliday: float
    TotalPaidHours: float


@dataclass
class MonthlyAdjustment:
    Name: str
    Type: str
    Value: float


@dataclass
class MonthlyWeekSummary:
    employees: list[MonthlyEmployee] = field(default_factory=list)
    employee_totals: list[MonthlyEmployeeTotal] = field(default_factory=list)
    adjustments: list[MonthlyAdjustment] = field(default_factory=list)
    start_date: str = ""
    end_date: str = ""
    non_agency_total: float = 0.0
    grouped_totals: dict[str, MonthlyEmployeeTotal] = field(default_factory=dict)


def _to_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() == "nan":
        return ""
    return text


def _parse_decimal(value: Any) -> float:
    text = _to_text(value).replace(",", "")
    if not text:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def _parse_int(value: Any) -> int:
    text = _to_text(value).replace(",", "")
    if not text:
        return 0
    try:
        return int(float(text))
    except ValueError:
        return 0


def _grouped_key(category: str) -> str:
    category = _to_text(category)
    if category.startswith("A-") and len(category) >= 9:
        return f"{category[:4]} {category[5:9]}"
    return category[:4] if len(category) >= 4 else category


def parse_monthly_week_file(file_obj: Any) -> MonthlyWeekSummary:
    from .payroll_service import _load_sheet  # reuse weekly reader

    table = _load_sheet(file_obj)
    rows = table.values.tolist()
    text_rows = [[_to_text(v) for v in row] for row in rows]
    if not text_rows:
        return MonthlyWeekSummary()

    out = MonthlyWeekSummary()
    out.start_date = _to_text(text_rows[0][4] if len(text_rows[0]) > 4 else "").removeprefix("D ").strip()
    out.end_date = _to_text(text_rows[0][7] if len(text_rows[0]) > 7 else "").removeprefix("D ").strip()

    r = 3  # row 4 in excel
    while r < len(text_rows):
        row = text_rows[r]
        if not _to_text(row[0] if len(row) > 0 else ""):
            break
        out.employees.append(
            MonthlyEmployee(
                Name=_to_text(row[0] if len(row) > 0 else ""),
                Category=_to_text(row[1] if len(row) > 1 else ""),
                SageNo=_parse_int(row[2] if len(row) > 2 else ""),
                BasicHours=_parse_decimal(row[3] if len(row) > 3 else ""),
                MonFriOvertime=_parse_decimal(row[4] if len(row) > 4 else ""),
                SatSunOvertime=_parse_decimal(row[5] if len(row) > 5 else ""),
                AnnualHoliday=_parse_decimal(row[6] if len(row) > 6 else ""),
                TotalPaidHours=_parse_decimal(row[7] if len(row) > 7 else ""),
            )
        )
        r += 1

    adjustments_header = -1
    for i in range(r, min(len(text_rows), r + 120)):
        if _to_text(text_rows[i][1] if len(text_rows[i]) > 1 else "").startswith("Adjustments"):
            adjustments_header = i
            break
    if adjustments_header >= 0:
        ar = adjustments_header + 2
        while ar < len(text_rows):
            name = _to_text(text_rows[ar][1] if len(text_rows[ar]) > 1 else "")
            if not name:
                break
            out.adjustments.append(
                MonthlyAdjustment(
                    Name=name,
                    Type=_to_text(text_rows[ar][2] if len(text_rows[ar]) > 2 else ""),
                    Value=_parse_decimal(text_rows[ar][3] if len(text_rows[ar]) > 3 else ""),
                )
            )
            ar += 1
        r = ar

    totals_header = -1
    for i in range(r, min(len(text_rows), r + 150)):
        if _to_text(text_rows[i][1] if len(text_rows[i]) > 1 else "") == "Category":
            totals_header = i
            break
    if totals_header >= 0:
        tr = totals_header + 1
        while tr < len(text_rows):
            cat = _to_text(text_rows[tr][1] if len(text_rows[tr]) > 1 else "")
            if not cat:
                break
            total = MonthlyEmployeeTotal(
                Category=cat,
                BasicHours=_parse_decimal(text_rows[tr][3] if len(text_rows[tr]) > 3 else ""),
                MonFriOvertime=_parse_decimal(text_rows[tr][4] if len(text_rows[tr]) > 4 else ""),
                SatSunOvertime=_parse_decimal(text_rows[tr][5] if len(text_rows[tr]) > 5 else ""),
                AnnualHoliday=_parse_decimal(text_rows[tr][6] if len(text_rows[tr]) > 6 else ""),
                TotalPaidHours=_parse_decimal(text_rows[tr][7] if len(text_rows[tr]) > 7 else ""),
            )
            out.employee_totals.append(total)
            tr += 1

    out.non_agency_total = sum(t.TotalPaidHours for t in out.employee_totals if not t.Category.startswith("A-"))
    grouped: dict[str, MonthlyEmployeeTotal] = {}
    for t in out.employee_totals:
        k = _grouped_key(t.Category)
        if k not in grouped:
            grouped[k] = MonthlyEmployeeTotal(k, 0.0, 0.0, 0.0, 0.0, 0.0)
        g = grouped[k]
        g.BasicHours += t.BasicHours
        g.MonFriOvertime += t.MonFriOvertime
        g.SatSunOvertime += t.SatSunOvertime
        g.AnnualHoliday += t.AnnualHoliday
        g.TotalPaidHours += t.TotalPaidHours
    out.grouped_totals = grouped
    return out


def parse_monthly_inputs(weekly_files: list[Any]) -> list[MonthlyWeekSummary]:
    return [parse_monthly_week_file(f) for f in weekly_files]


def monthly_summaries_to_json(summaries: list[MonthlyWeekSummary]) -> list[dict[str, Any]]:
    return [asdict(s) for s in summaries]


def monthly_summaries_from_json(data: list[dict[str, Any]]) -> list[MonthlyWeekSummary]:
    out: list[MonthlyWeekSummary] = []
    for d in data:
        s = MonthlyWeekSummary(
            employees=[MonthlyEmployee(**e) for e in d.get("employees", [])],
            employee_totals=[MonthlyEmployeeTotal(**t) for t in d.get("employee_totals", [])],
            adjustments=[MonthlyAdjustment(**a) for a in d.get("adjustments", [])],
            start_date=str(d.get("start_date", "")),
            end_date=str(d.get("end_date", "")),
            non_agency_total=float(d.get("non_agency_total", 0.0)),
        )
        s.grouped_totals = {}
        for k, v in (d.get("grouped_totals") or {}).items():
            if isinstance(v, dict):
                s.grouped_totals[str(k)] = MonthlyEmployeeTotal(**v)
        out.append(s)
    return out


def _write_header(ws, r: int, c: int = 1) -> None:
    ws.cell(r, c, "Name")
    ws.cell(r, c + 1, "Category")
    ws.cell(r, c + 2, "SageNo")
    ws.cell(r, c + 3, "BasicHours")
    ws.cell(r, c + 4, "MonFriOvertime")
    ws.cell(r, c + 5, "SatSunOvertime")
    ws.cell(r, c + 6, "AnnualHoliday")
    ws.cell(r, c + 7, "TotalPaidHours")


def _write_total_header(ws, r: int) -> None:
    ws.cell(r, 2, "Category")
    ws.cell(r, 4, "BasicHours")
    ws.cell(r, 5, "MonFriOvertime")
    ws.cell(r, 6, "SatSunOvertime")
    ws.cell(r, 7, "AnnualHoliday")
    ws.cell(r, 8, "TotalPaidHours")


def build_monthly_excel_bytes(
    week_summaries: list[MonthlyWeekSummary],
    non_hourly_names: set[str] | None = None,
) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    grouped_from_weeks: list[dict[str, MonthlyEmployeeTotal]] = []
    for i, s in enumerate(week_summaries, start=1):
        ws = wb.create_sheet(f"Week{i}")
        ws.cell(1, 1, f"Week {i}")
        ws.cell(1, 4, "Start Date")
        ws.cell(1, 5, f"D {s.start_date}")
        ws.cell(1, 7, "End Date")
        ws.cell(1, 8, f"D {s.end_date}")
        _write_header(ws, 3)

        r = 4
        for e in s.employees:
            ws.cell(r, 1, e.Name)
            ws.cell(r, 2, e.Category)
            ws.cell(r, 3, e.SageNo)
            ws.cell(r, 4, e.BasicHours)
            ws.cell(r, 5, e.MonFriOvertime)
            ws.cell(r, 6, e.SatSunOvertime)
            ws.cell(r, 7, e.AnnualHoliday)
            ws.cell(r, 8, e.TotalPaidHours)
            r += 1

        r += 2
        ws.cell(r, 2, f"Adjustments - {len(s.adjustments)}")
        r += 1
        ws.cell(r, 2, "Name")
        ws.cell(r, 3, "Type")
        ws.cell(r, 4, "Value")
        r += 1
        for a in s.adjustments:
            ws.cell(r, 2, a.Name)
            ws.cell(r, 3, a.Type)
            ws.cell(r, 4, a.Value)
            r += 1

        r += 2
        _write_total_header(ws, r)
        r += 1
        non_agency_bands = [0.0, 0.0, 0.0, 0.0, 0.0]
        for t in s.employee_totals:
            ws.cell(r, 2, t.Category)
            ws.cell(r, 4, t.BasicHours)
            ws.cell(r, 5, t.MonFriOvertime)
            ws.cell(r, 6, t.SatSunOvertime)
            ws.cell(r, 7, t.AnnualHoliday)
            ws.cell(r, 8, t.TotalPaidHours)
            if not t.Category.startswith("A-"):
                non_agency_bands[0] += t.BasicHours
                non_agency_bands[1] += t.MonFriOvertime
                non_agency_bands[2] += t.SatSunOvertime
                non_agency_bands[3] += t.AnnualHoliday
                non_agency_bands[4] += t.TotalPaidHours
            r += 1

        r += 1
        for k, v in s.grouped_totals.items():
            ws.cell(r, 2, k)
            ws.cell(r, 4, v.BasicHours)
            ws.cell(r, 5, v.MonFriOvertime)
            ws.cell(r, 6, v.SatSunOvertime)
            ws.cell(r, 7, v.AnnualHoliday)
            ws.cell(r, 8, v.TotalPaidHours)
            r += 1

        r += 2
        ws.cell(r, 4, non_agency_bands[0])
        ws.cell(r, 5, non_agency_bands[1])
        ws.cell(r, 6, non_agency_bands[2])
        ws.cell(r, 7, non_agency_bands[3])
        ws.cell(r, 8, non_agency_bands[4])
        grouped_from_weeks.append(s.grouped_totals)

    merged_employees: dict[str, MonthlyEmployee] = {}
    merged_totals: dict[str, MonthlyEmployeeTotal] = {}
    for s in week_summaries:
        for e in s.employees:
            cur = merged_employees.get(e.Name)
            if cur is None:
                merged_employees[e.Name] = MonthlyEmployee(**e.__dict__)
                continue
            cur.BasicHours += e.BasicHours
            cur.MonFriOvertime += e.MonFriOvertime
            cur.SatSunOvertime += e.SatSunOvertime
            cur.AnnualHoliday += e.AnnualHoliday
            cur.TotalPaidHours += e.TotalPaidHours
            cur.SageNo = e.SageNo
            cur.Category = e.Category

        for t in s.employee_totals:
            cur_t = merged_totals.get(t.Category)
            if cur_t is None:
                merged_totals[t.Category] = MonthlyEmployeeTotal(**t.__dict__)
                continue
            cur_t.BasicHours += t.BasicHours
            cur_t.MonFriOvertime += t.MonFriOvertime
            cur_t.SatSunOvertime += t.SatSunOvertime
            cur_t.AnnualHoliday += t.AnnualHoliday
            cur_t.TotalPaidHours += t.TotalPaidHours

    non_hourly_names = {n.strip().upper() for n in (non_hourly_names or set()) if n and n.strip()}
    for e in merged_employees.values():
        if e.Name.strip().upper() in non_hourly_names:
            e.IsHourly = False
        if not e.IsHourly and e.Category.strip().startswith("A-"):
            raise ValueError(f"Agency employee cannot be non-hourly: {e.Name}")

    ws = wb.create_sheet("Summary")
    _write_header(ws, 3)
    r = 4
    for e in merged_employees.values():
        ws.cell(r, 1, e.Name)
        ws.cell(r, 2, e.Category)
        ws.cell(r, 3, e.SageNo)
        ws.cell(r, 4, e.BasicHours)
        ws.cell(r, 5, e.MonFriOvertime)
        ws.cell(r, 6, e.SatSunOvertime)
        ws.cell(r, 7, e.AnnualHoliday)
        ws.cell(r, 8, e.TotalPaidHours)
        r += 1

    r += 2
    _write_total_header(ws, r)
    r += 1
    non_agency_summary = [0.0, 0.0, 0.0, 0.0, 0.0]
    for t in merged_totals.values():
        ws.cell(r, 2, t.Category)
        ws.cell(r, 4, t.BasicHours)
        ws.cell(r, 5, t.MonFriOvertime)
        ws.cell(r, 6, t.SatSunOvertime)
        ws.cell(r, 7, t.AnnualHoliday)
        ws.cell(r, 8, t.TotalPaidHours)
        if not t.Category.startswith("A-"):
            non_agency_summary[0] += t.BasicHours
            non_agency_summary[1] += t.MonFriOvertime
            non_agency_summary[2] += t.SatSunOvertime
            non_agency_summary[3] += t.AnnualHoliday
            non_agency_summary[4] += t.TotalPaidHours
        non_hourly = [e for e in merged_employees.values() if not e.IsHourly and e.Category.upper() == t.Category.upper()]
        if non_hourly:
            ws.cell(r + 1, 2, f"{t.Category} non-hourly hours")
            ws.cell(r + 1, 4, -sum(e.BasicHours for e in non_hourly))
            ws.cell(r + 1, 5, -sum(e.MonFriOvertime for e in non_hourly))
            ws.cell(r + 1, 6, -sum(e.SatSunOvertime for e in non_hourly))
            ws.cell(r + 1, 7, -sum(e.AnnualHoliday for e in non_hourly))
            ws.cell(r + 1, 8, -sum(e.TotalPaidHours for e in non_hourly))
            r += 2
        else:
            r += 1

    r += 1
    agg_grouped: dict[str, MonthlyEmployeeTotal] = {}
    for m in grouped_from_weeks:
        for k, t in m.items():
            if k not in agg_grouped:
                agg_grouped[k] = MonthlyEmployeeTotal(k, 0.0, 0.0, 0.0, 0.0, 0.0)
            a = agg_grouped[k]
            a.BasicHours += t.BasicHours
            a.MonFriOvertime += t.MonFriOvertime
            a.SatSunOvertime += t.SatSunOvertime
            a.AnnualHoliday += t.AnnualHoliday
            a.TotalPaidHours += t.TotalPaidHours
    for t in agg_grouped.values():
        ws.cell(r, 2, t.Category)
        ws.cell(r, 4, t.BasicHours)
        ws.cell(r, 5, t.MonFriOvertime)
        ws.cell(r, 6, t.SatSunOvertime)
        ws.cell(r, 7, t.AnnualHoliday)
        ws.cell(r, 8, t.TotalPaidHours)
        r += 1

    r += 2
    ws.cell(r, 4, non_agency_summary[0])
    ws.cell(r, 5, non_agency_summary[1])
    ws.cell(r, 6, non_agency_summary[2])
    ws.cell(r, 7, non_agency_summary[3])
    ws.cell(r, 8, non_agency_summary[4])

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()

from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from typing import Any

import pandas as pd


AGENCY_CATEGORIES = {
    "A-EL CLNR",
    "A-EL DPCH",
    "A-EL PKNG HIGH_RISK",
    "A-EL PKNG SLEEVING",
    "A-EL PROD BELT",
    "A-EL PROD FORMING",
    "A-EL PROD FRYING",
    "A-EL PROD PREPARATION",
    "A-EL TECHNICAL",
    "A-PM PKNG HIGH_RISK",
    "A-RS PKNG SLEEVING",
    "A-RS PROD BELT",
    "A-RS PROD FRYING",
}


@dataclass
class PayrollResult:
    rows: list[dict[str, Any]]
    agency_rows: list[dict[str, Any]]
    gazebo_rows: list[dict[str, Any]]
    total_paid_hours_non_agency: float


def _normalize_header(text: str) -> str:
    return "".join(ch.lower() for ch in (text or "") if ch.isalnum())


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


def _parse_int(value: Any) -> int | None:
    text = _to_text(value).replace(",", "")
    if not text:
        return None
    try:
        return int(float(text))
    except ValueError:
        return None


def _row_contains_date_range(row: list[str]) -> bool:
    return any("date range" in cell.lower() for cell in row)


def _is_category_row(row: list[str], pay_id_col: int) -> bool:
    pay = row[pay_id_col] if pay_id_col < len(row) else ""
    if _parse_int(pay) is not None:
        return False
    c0 = row[0] if row else ""
    return bool(c0 and not c0.lower().startswith("total for"))


def _load_sheet(file_obj: Any) -> pd.DataFrame:
    file_obj.seek(0)
    return pd.read_excel(file_obj, sheet_name=0, header=None, dtype=str)


def parse_employee_hours(file_obj: Any) -> list[dict[str, Any]]:
    df = _load_sheet(file_obj)
    rows = [[_to_text(v) for v in rec] for rec in df.values.tolist()]
    if not rows:
        return []

    header_row = -1
    pay_id_col = -1
    scan_limit = min(len(rows), 40)
    for r in range(scan_limit):
        for c, cell in enumerate(rows[r]):
            if _normalize_header(cell) == "payid":
                header_row = r
                pay_id_col = c
                break
        if header_row >= 0:
            break

    if header_row < 0 or pay_id_col < 0:
        return _parse_employee_legacy(rows)

    name_col = pay_id_col - 1 if pay_id_col > 0 else 0
    basic_col = pay_id_col + 1
    mon_fri_col = pay_id_col + 2
    sat_sun_col = pay_id_col + 3
    annual_col = pay_id_col + 4

    category = ""
    result: list[dict[str, Any]] = []
    for r in range(header_row + 1, len(rows)):
        row = rows[r]
        if _row_contains_date_range(row):
            break
        first = row[0] if row else ""
        if first.lower().startswith("total for"):
            continue

        pay_text = row[pay_id_col] if pay_id_col < len(row) else ""
        name_text = row[name_col] if name_col < len(row) else ""
        pay_id = _parse_int(pay_text)
        if pay_id is None:
            if _is_category_row(row, pay_id_col):
                category = first
            continue

        if not name_text or name_text.upper() == "U EMPLOYEE" or "total for" in name_text.lower():
            continue

        basic_hours = _parse_decimal(row[basic_col] if basic_col < len(row) else "")
        mon_fri_ot = _parse_decimal(row[mon_fri_col] if mon_fri_col < len(row) else "")
        sat_sun_ot = _parse_decimal(row[sat_sun_col] if sat_sun_col < len(row) else "")
        annual_holiday = _parse_decimal(row[annual_col] if annual_col < len(row) else "")
        total_paid = basic_hours + mon_fri_ot + sat_sun_ot + annual_holiday

        result.append(
            {
                "Name": name_text.upper(),
                "Category": category,
                "SageNo": pay_id,
                "BasicHours": basic_hours,
                "MonFriOvertime": mon_fri_ot,
                "SatSunOvertime": sat_sun_ot,
                "AnnualHoliday": annual_holiday,
                "TotalPaidHours": total_paid,
            }
        )

    return result


def _parse_employee_legacy(rows: list[list[str]]) -> list[dict[str, Any]]:
    category = ""
    expects_category = True
    result: list[dict[str, Any]] = []
    idx = 0
    while idx < 250 and idx + 6 < len(rows):
        row = rows[6 + idx]
        col_a = row[0] if row else ""
        if col_a:
            if expects_category:
                if col_a.lower() != "default":
                    category = col_a
                expects_category = False
            else:
                if any("date range" in cell.lower() for cell in row):
                    break
                name = _to_text(row[0]).upper()
                if name and name != "U EMPLOYEE":
                    basic = _parse_decimal(row[2] if len(row) > 2 else "")
                    mon_fri = _parse_decimal(row[3] if len(row) > 3 else "")
                    sat_sun = _parse_decimal(row[6] if len(row) > 6 else "")
                    annual = _parse_decimal(row[7] if len(row) > 7 else "")
                    basic_after_legacy = basic - annual
                    result.append(
                        {
                            "Name": name,
                            "Category": category,
                            "SageNo": _parse_int(row[1] if len(row) > 1 else "") or 0,
                            "BasicHours": basic_after_legacy,
                            "MonFriOvertime": mon_fri,
                            "SatSunOvertime": sat_sun,
                            "AnnualHoliday": annual,
                            "TotalPaidHours": basic_after_legacy + mon_fri + sat_sun + annual,
                        }
                    )
        else:
            expects_category = True
            idx += 1
        idx += 1
    return result


def _find_header_index(headers: list[str], *candidates: str) -> int:
    for idx, cell in enumerate(headers):
        n = _normalize_header(cell)
        if any(n == c for c in candidates):
            return idx
    return -1


def parse_contracted_hours(file_obj: Any) -> tuple[dict[int, float], dict[str, float]]:
    df = _load_sheet(file_obj)
    rows = [[_to_text(v) for v in rec] for rec in df.values.tolist()]
    if not rows:
        return {}, {}

    header_row = -1
    payroll_col = -1
    contract_col = -1
    name_col = -1
    for r, row in enumerate(rows[:80]):
        payroll_col = _find_header_index(row, "payrollnumber", "payrollno", "payroll")
        contract_col = _find_header_index(row, "contracthrs", "contracthours", "contracthr", "contractedhours")
        name_col = _find_header_index(row, "clockname", "name", "employeename")
        if payroll_col >= 0 and contract_col >= 0:
            header_row = r
            break

    if header_row < 0:
        return {}, {}

    by_payroll: dict[int, float] = {}
    by_name: dict[str, float] = {}
    for row in rows[header_row + 1 :]:
        hours = _parse_decimal(row[contract_col] if contract_col < len(row) else "")
        if hours == 0.0:
            continue
        pay_no = _parse_int(row[payroll_col] if payroll_col < len(row) else "")
        if pay_no is not None:
            by_payroll[pay_no] = hours
        if name_col >= 0 and name_col < len(row):
            name = _to_text(row[name_col]).upper()
            if name:
                by_name[name] = hours

    return by_payroll, by_name


def calculate_payroll(employee_rows: list[dict[str, Any]], contracted_file_obj: Any) -> PayrollResult:
    by_payroll, by_name = parse_contracted_hours(contracted_file_obj)

    for row in employee_rows:
        contracted = by_payroll.get(int(row["SageNo"]))
        if contracted is None:
            contracted = by_name.get(str(row["Name"]).upper(), 0.0)
            row["ContractHourMatch"] = "Yes" if contracted else "No"
        else:
            row["ContractHourMatch"] = "Yes"
        row["ContractedHours"] = float(contracted)
        row["Overtime"] = float(row["TotalPaidHours"]) - float(row["ContractedHours"])

    agency_rows = [r for r in employee_rows if str(r["Category"]).strip().upper() in AGENCY_CATEGORIES]
    gazebo_rows = [r for r in employee_rows if str(r["Category"]).strip().upper() not in AGENCY_CATEGORIES]
    total_non_agency = sum(float(r["TotalPaidHours"]) for r in employee_rows if not str(r["Category"]).startswith("A-"))

    return PayrollResult(
        rows=employee_rows,
        agency_rows=agency_rows,
        gazebo_rows=gazebo_rows,
        total_paid_hours_non_agency=round(total_non_agency, 2),
    )


def build_excel_bytes(result: PayrollResult) -> bytes:
    all_df = pd.DataFrame(result.rows)
    agency_df = pd.DataFrame(result.agency_rows)
    gazebo_df = pd.DataFrame(result.gazebo_rows)

    if all_df.empty:
        analysis_df = pd.DataFrame(columns=["Category", "BasicHours", "MonFriOvertime", "SatSunOvertime", "AnnualHoliday", "TotalPaidHours"])
    else:
        analysis_df = (
            all_df.groupby("Category", dropna=False)[["BasicHours", "MonFriOvertime", "SatSunOvertime", "AnnualHoliday", "TotalPaidHours"]]
            .sum()
            .reset_index()
        )

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        all_df.to_excel(writer, sheet_name="All Data", index=False)
        agency_df.to_excel(writer, sheet_name="Agency Employee", index=False)
        gazebo_df.to_excel(writer, sheet_name="Gazebo Employee", index=False)
        analysis_df.to_excel(writer, sheet_name="Analysis", index=False)
    output.seek(0)
    return output.getvalue()

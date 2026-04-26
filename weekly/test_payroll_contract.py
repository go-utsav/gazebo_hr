"""Contract hours parsing: Clockrite report layout vs tabular XLS."""
from __future__ import annotations

import unittest
from pathlib import Path

from .payroll_service import calculate_payroll, parse_contracted_hours, parse_employee_hours, total_paid_hours_from_rows

class TotalPaidHoursFromRowsTest(unittest.TestCase):
    def test_sums_including_a_prefix(self) -> None:
        rows = [
            {"TotalPaidHours": 4.0, "Category": "A-EL PROD"},
            {"TotalPaidHours": 3.5, "Category": "D-STAFF"},
        ]
        self.assertEqual(total_paid_hours_from_rows(rows), 7.5)


_DATA = Path(__file__).resolve().parent.parent / "data"
_CLOCKRITE = _DATA / "Employee contract hours - clockrite.xls"
_EMPLOYEE = _DATA / "dgross_paysummary2.xls"


@unittest.skipUnless(_CLOCKRITE.is_file() and _EMPLOYEE.is_file(), "data/*.xls fixtures not in repo")
class ClockriteContractParseTest(unittest.TestCase):
    def test_parse_clockrite_export_maps_payroll_number_to_contract_hours(self) -> None:
        with _CLOCKRITE.open("rb") as f:
            by_payroll, _ = parse_contracted_hours(f)
        self.assertGreater(
            len(by_payroll),
            100,
            "Expected Clockrite 'Employee Details' report to yield many payroll keys",
        )
        # Known block from sample file: Payroll Number 1026, Contract Hrs 0
        self.assertIn(1026, by_payroll)
        self.assertEqual(by_payroll[1026], 0.0)

    def test_join_employee_pay_id_to_contract_payroll_number(self) -> None:
        with _EMPLOYEE.open("rb") as ef, _CLOCKRITE.open("rb") as cf:
            employees = parse_employee_hours(ef)
            result = calculate_payroll(employees, cf)
        matched = sum(1 for r in result.rows if r["ContractHourMatch"] == "Yes")
        # Sample data: 189/190 Pay IDs exist as Payroll Number in the contract file
        self.assertGreaterEqual(matched, 180)
        self.assertLess(matched, len(result.rows) + 1)

"""Tests for the parts of update_territory_sheet.py that decide what reaches (or leaves) the Sheet.
Run: python3 -m unittest discover -s sync -p "test_*.py"
Uses only made-up rows and a tiny temporary xlsx, never the real source files."""
import sys
import tempfile
import unittest
from datetime import date
from unittest import mock
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
import update_territory_sheet as u  # noqa: E402


def rtr_month(sheet_name, day):
    return {"sheetName": sheet_name, "label": sheet_name, "rowCount": 0, "stores": [],
            "cmRollup": {}, "daily": {day: {"ga": 1}}}


def tables_for(key):
    return {f"{key}-meta": {"headers": ["field", "value"], "rows": []}}


class MonthYear(unittest.TestCase):
    def test_sheet_without_year_takes_the_year_of_its_data(self):
        d = u.parse_sheet_date("Oct")
        self.assertEqual(u.month_year({"daily": {"20261031": {}}}, d, 2027), 2026)

    def test_year_in_sheet_name_wins(self):
        d = u.parse_sheet_date("Sep.26")
        self.assertEqual(u.month_year({"daily": {}}, d, 2027), 2026)


class BuildPayloadRetention(unittest.TestCase):
    def test_a_q4_sheet_run_in_january_keeps_real_months(self):
        existing_keys = [f"rtr-2026-{m:02d}" for m in range(6, 12)]  # Jun..Nov 2026
        existing_tables = {}
        for k in existing_keys:
            existing_tables.update(tables_for(k))
        january = mock.Mock(wraps=date)
        january.today.return_value = date(2027, 1, 5)
        with mock.patch.object(u, "date", january):
            payload = u.build_payload(
                [rtr_month("Dec", "20261231")], [], [],
                existing_tables=existing_tables,
                existing_meta={"rtr_months": existing_keys, "sev_months": []},
            )
        self.assertEqual(payload["meta"]["rtr_months"],
                         [f"rtr-2026-{m:02d}" for m in range(7, 13)])
        self.assertEqual(payload["deleteTabs"], ["rtr-2026-06-meta"])
        self.assertFalse(any(t.startswith("rtr-2027") for t in payload["tables"]))


class KpiSetupTargets(unittest.TestCase):
    def test_reads_every_month_block_in_one_header_row(self):
        rows = [
            [None, "Target GA (Oct)", None, None, "Target GA (Nov)", None, None],
            ["COM_Name", "Mass + Migrant", "Tourist", "Total GA", "Mass + Migrant", "Tourist", "Total GA"],
            ["Com A", 10, 5, 15, 20, 0, 20],
            ["Com B", 4, 0, 4, 6, 1, 7],
        ]
        got = u.parse_kpi_setup_targets(rows, name_col="com_name")
        self.assertEqual(got["oct"]["Com A"]["total"], 15)
        self.assertEqual(got["nov"]["Com A"]["total"], 20)
        self.assertEqual(got["nov"]["Com B"]["tourist"], 1)


class TiersWithoutTourist(unittest.TestCase):
    def test_tourist_rows_count_in_tiers_but_not_in_tiers_mm(self):
        header = ["PARTNER_CODE", "PARTNER_NAME", "GA", "GROUP_SIM", "D_CLUSTER", "TM_KEY_DAY", "AP1D_300UP"]
        rows = [header,
                ["S1", "Shop", 5, "MASS", "Pak Kret", "20261101", 2],
                ["S1", "Shop", 9, "TOURIST", "Pak Kret", "20261101", 7]]
        got = u.aggregate_rtr_sheet(rows, {}, {"S1": "Cm A"}, {})
        s = got["stores"][0]
        self.assertEqual(s["tiers"][7], 9)
        self.assertEqual(s["tiersMM"][7], 2)
        table = u.build_rtr_month_tables({"sheetName": "Nov", "label": "Nov 2026", "rowCount": 2,
                                          "stores": got["stores"], "cmRollup": {}})
        hdr = table["stores"]["headers"]
        self.assertEqual(table["stores"]["rows"][0][hdr.index("tierMM_300+")], 2)
        self.assertEqual(len(hdr), len(table["stores"]["rows"][0]))


class RoutingNewStores(unittest.TestCase):
    def test_a_store_only_the_routing_sheet_lists_is_added_with_both_codes(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "routing.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "RTR_Route"
            ws.append(["Dtac code", "True code", "Retailer Name", "Cluster", "COM", "CM", "อำเภอ"])
            ws.append(["D1", "T1", "Old Shop", "Pak Kret", "Com A", "Cm A", "ปากเกร็ด"])
            ws.append(["D2", "T2", "New Shop", "Pak Kret", "Com B", "Cm B", "ปากเกร็ด"])
            ws.append(["D3", "T3", "Other Area", "Bang Bua Thong", "Com C", "Cm C", ""])
            wb.save(path)

            code_to_canonical = {"D1": "D1", "T1": "D1"}
            new_stores = {}
            _, to_com, to_cm = u.extract_routing_overrides(path, code_to_canonical, new_stores)

        self.assertEqual(list(new_stores), ["D2"])
        self.assertEqual(new_stores["D2"]["name"], "New Shop")
        self.assertEqual(code_to_canonical["T2"], "D2")
        self.assertEqual(to_com["D2"], "Com B")
        self.assertEqual(to_cm["D2"], "Cm B")
        self.assertNotIn("D3", code_to_canonical)


if __name__ == "__main__":
    unittest.main()

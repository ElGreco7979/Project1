"""Basic tests: python -m unittest test_cy_vat_working.py"""
import datetime as dt
import tempfile
import unittest
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

import cy_vat_working as cvw


class TestHelpers(unittest.TestCase):
    def test_quarter_dates(self):
        self.assertEqual(cvw.quarter_dates(2026, 1), (dt.date(2026, 1, 1), dt.date(2026, 3, 31)))
        self.assertEqual(cvw.quarter_dates(2026, 4), (dt.date(2026, 10, 1), dt.date(2026, 12, 31)))

    def test_parse_ecb_csv(self):
        text = ("KEY,FREQ,CURRENCY,CURRENCY_DENOM,EXR_TYPE,EXR_SUFFIX,TIME_PERIOD,OBS_VALUE\n"
                "EXR.D.USD.EUR.SP00.A,D,USD,EUR,SP00,A,2026-07-01,1.1702\n"
                "EXR.D.GBP.EUR.SP00.A,D,GBP,EUR,SP00,A,2026-07-01,0.8601\n")
        self.assertEqual(cvw.parse_ecb_csv(text), {dt.date(2026, 7, 1): {"USD": 1.1702, "GBP": 0.8601}})

    def test_to_number(self):
        self.assertEqual(cvw.to_number("(1,250.50)"), -1250.5)
        self.assertIsNone(cvw.to_number(""))


class TestWorkflow(unittest.TestCase):
    def test_new_and_import(self):
        with tempfile.TemporaryDirectory() as d:
            wbp, glp = Path(d) / "w.xlsx", Path(d) / "gl.csv"
            cvw.main(["new", str(wbp), "--year", "2026", "--quarter", "3"])
            pd.DataFrame([{"Date": "05/08/2026", "Doc No": "1", "Account": "4000", "Account Name": "Sales",
                           "Description": "x", "Counterparty": "C1", "Counterparty Name": "A",
                           "Currency": "EUR", "Amount": "-100"}]).to_csv(glp, index=False)
            cvw.main(["import-gl", str(wbp), str(glp)])
            wb = load_workbook(wbp)
            self.assertEqual(wb["GL_Input"]["C5"].value, "4000")
            self.assertEqual(wb["GL_Input"]["A5"].value.date(), dt.date(2026, 8, 5))
            self.assertEqual(wb["Account_Mapping"]["A5"].value, "4000")   # examples removed
            self.assertEqual(wb["Counterparties"]["A5"].value, "C1")


if __name__ == "__main__":
    unittest.main()

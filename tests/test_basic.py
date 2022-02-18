import unittest

from openpyxl import Workbook, load_workbook
from src.DataRow import DateRow


class BasicTestCase(unittest.TestCase):
    def test_init(self):
        self.assertIsInstance(DateRow(None), DateRow)

    def test_generate(self):
        dr = DateRow(None)
        self.assertGreaterEqual(len([x for x in dr.year_iterator()]), 356)

    def test_add_row(self):
        wb = Workbook()
        dr = DateRow(wb.active)
        self.assertTrue(dr)
        self.assertTrue(dr.add_row())
        wb.save('test.xlsx')

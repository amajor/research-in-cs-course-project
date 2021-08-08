import datetime
import unittest
from spreadsheets.excel_order_two import WORKBOOK, ORDER_TWO_CELLS
from spreadsheets.methods import get_po_number, get_bill_to_name, get_ship_to_name, get_ship_date, get_total_quantity, \
    get_total_cost


class ExcelOrderTwo(unittest.TestCase):
    def test_get_po_number(self):
        expected = "1234567"
        actual = get_po_number(WORKBOOK, ORDER_TWO_CELLS)
        self.assertEqual(expected, actual)

    def test_get_bill_to_name(self):
        expected = "[cell not mapped]"
        actual = get_bill_to_name(WORKBOOK, ORDER_TWO_CELLS)
        self.assertEqual(expected, actual)

    def test_get_ship_to_name(self):
        expected = "XYZ Back Door"
        actual = get_ship_to_name(WORKBOOK, ORDER_TWO_CELLS)
        self.assertEqual(expected, actual)

    def test_get_ship_date(self):
        expected = datetime.datetime(2021, 9, 1, 0, 0)  # 09/01/21
        actual = get_ship_date(WORKBOOK, ORDER_TWO_CELLS)
        self.assertEqual(expected, actual)

    def test_get_total_quantity(self):
        expected = 23
        actual = get_total_quantity(WORKBOOK, ORDER_TWO_CELLS)
        self.assertEqual(expected, actual)

    def test_get_total_cost(self):
        expected = 478.00
        actual = get_total_cost(WORKBOOK, ORDER_TWO_CELLS)
        self.assertEqual(expected, actual)


if __name__ == '__main__':
    unittest.main()

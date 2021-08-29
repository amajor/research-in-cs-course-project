import datetime
import unittest
from parameterized import parameterized
from spreadsheets.excel_order_one import WORKBOOK, ORDER_ONE_CELLS, ITEM_TABLE
from spreadsheets.methods import *


class ExcelOrderOne(unittest.TestCase):
    def test_get_po_number(self):
        expected = "1234567"
        actual = get_po_number(WORKBOOK, ORDER_ONE_CELLS)
        self.assertEqual(expected, actual)

    def test_get_bill_to_name(self):
        expected = "XYZ Company"
        actual = get_bill_to_name(WORKBOOK, ORDER_ONE_CELLS)
        self.assertEqual(expected, actual)

    def test_get_ship_to_name(self):
        expected = "XYZ Back Door"
        actual = get_ship_to_name(WORKBOOK, ORDER_ONE_CELLS)
        self.assertEqual(expected, actual)

    def test_get_ship_date(self):
        expected = datetime.datetime(2021, 9, 1, 0, 0)  # 09/01/21
        actual = get_ship_date(WORKBOOK, ORDER_ONE_CELLS)
        self.assertEqual(expected, actual)

    def test_get_total_quantity(self):
        expected = 23
        actual = get_total_quantity(WORKBOOK, ORDER_ONE_CELLS)
        self.assertEqual(expected, actual)

    def test_get_total_cost(self):
        expected = 478.00
        actual = get_total_cost(WORKBOOK, ORDER_ONE_CELLS)
        self.assertEqual(expected, actual)

    @parameterized.expand([
        ["Row 1 gives line '1'", ITEM_TABLE, 0, 1],
        ["Row 2 gives line '2'", ITEM_TABLE, 1, 2],
        ["Row 3 gives line '3'", ITEM_TABLE, 2, 3],
        ["Row 4 gives line '4'", ITEM_TABLE, 3, 4]
    ])
    def test_get_item_line_number(self, name, table, row, expected):
        actual = get_item_line_number(table, row)
        self.assertEqual(expected, actual)

    @parameterized.expand([
        ["Row 1 gives quantity 7", ITEM_TABLE, 0, 7],
        ["Row 2 gives quantity 6", ITEM_TABLE, 1, 6],
        ["Row 3 gives quantity 8", ITEM_TABLE, 2, 8],
        ["Row 4 gives quantity 2", ITEM_TABLE, 3, 2]
    ])
    def test_get_item_quantity(self, name, table, row, expected):
        actual = get_item_quantity(table, row)
        self.assertEqual(expected, actual)

    @parameterized.expand([
        ["Row 1 gives 'CASE'", ITEM_TABLE, 0, 'CASE'],
        ["Row 2 gives 'CASE'", ITEM_TABLE, 1, 'CASE'],
        ["Row 3 gives 'CASE'", ITEM_TABLE, 2, 'CASE'],
        ["Row 4 gives 'CASE'", ITEM_TABLE, 3, 'CASE']
    ])
    def test_get_item_unit_of_measure(self, name, table, row, expected):
        actual = get_item_unit_of_measure(table, row)
        self.assertEqual(expected, actual)

    @parameterized.expand([
        ["Row 1 gives item number '123'", ITEM_TABLE, 0, '123'],
        ["Row 2 gives item number '124'", ITEM_TABLE, 1, '122'],
        ["Row 3 gives item number '123'", ITEM_TABLE, 2, '124'],
        ["Row 4 gives item number '123'", ITEM_TABLE, 3, '143']
    ])
    def test_get_item_number(self, name, table, row, expected):
        actual = get_item_number(table, row)
        self.assertEqual(expected, actual)

    @parameterized.expand([
        ["Row 1 gives the description", ITEM_TABLE, 0, 'Green Balls\n16" Diameter'],
        ["Row 2 gives the description", ITEM_TABLE, 1, 'Red Balls\n12" Diameter'],
        ["Row 3 gives the description", ITEM_TABLE, 2, 'Blue Balls\n24" Diameter'],
        ["Row 4 gives the description", ITEM_TABLE, 3, 'Orange Balls\n6" Diameter']
    ])
    def test_get_item_description(self, name, table, row, expected):
        actual = get_item_description(table, row)
        self.assertEqual(expected, actual)

    @parameterized.expand([
        ["Row 1 gives 20.00", ITEM_TABLE, 0, 20.00],
        ["Row 2 gives 18.00", ITEM_TABLE, 1, 18.00],
        ["Row 3 gives 25.00", ITEM_TABLE, 2, 25.00],
        ["Row 4 gives 15.00", ITEM_TABLE, 3, 15.00]
    ])
    def test_get_item_unit_price(self, name, table, row, expected):
        actual = get_item_unit_price(table, row)
        self.assertEqual(expected, actual)

    @parameterized.expand([
        ["Row 1 gives 140.00", ITEM_TABLE, 0, 140.00],
        ["Row 2 gives 108.00", ITEM_TABLE, 1, 108.00],
        ["Row 3 gives 200.00", ITEM_TABLE, 2, 200.00],
        ["Row 4 gives 30.00", ITEM_TABLE, 3, 30.00]
    ])
    def test_get_item_cost(self, name, table, row, expected):
        actual = get_item_cost(table, row)
        self.assertEqual(expected, actual)


if __name__ == '__main__':
    unittest.main()

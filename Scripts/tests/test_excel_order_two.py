import datetime
import unittest
from parameterized import parameterized
from spreadsheets.excel_order_two import WORKBOOK, ORDER_TWO_CELLS, ITEM_TABLE
from spreadsheets.methods import get_po_number, get_bill_to_name, get_ship_to_name, get_ship_date, get_total_quantity, \
    get_total_cost, get_item_line_number, get_item_quantity, get_item_unit_of_measure, get_item_number, \
    get_item_description, get_item_unit_price, get_item_cost


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

    @parameterized.expand([
        ["Row 1 gives [not available]", ITEM_TABLE, 0, "[not available]"],
        ["Row 2 gives [not available]", ITEM_TABLE, 1, "[not available]"],
        ["Row 3 gives [not available]", ITEM_TABLE, 2, "[not available]"],
        ["Row 4 gives [not available]", ITEM_TABLE, 3, "[not available]"]
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

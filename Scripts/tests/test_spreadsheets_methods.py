import datetime
import unittest
import pandas as pd
from parameterized import parameterized
from spreadsheets.methods import get_column_value, get_row_value, get_po_number, get_total_cost, get_total_quantity, \
    get_ship_date, get_ship_to_name, get_bill_to_name, get_cell_value

FILE_PATH = '../documents/ORDERS.xlsx'
SHEET_NAME = 'Order Sample 1'

# The workbook to pull values from
WORKBOOK = pd.read_excel(
    FILE_PATH,
    sheet_name=SHEET_NAME,
    header=None
)

# Mapping of the values to cells
CELLS = {
    "poNum": "B1",
    "billTo": "B4",
    "shipTo": "B3",
    "shipDate": "B2",
    "totalQty": "B12",
    "totalCost": "G12"
}


class SpreadsheetsMethods(unittest.TestCase):
    @parameterized.expand([
        ["B1 gives 1", "B1", 1],
        ["B2 gives 1", "B2", 1],
        ["A1 gives 0", "A1", 0],
        ["Z9 gives 25", "Z9", 25]
    ])
    def test_get_column_value(self, name, cell, expected):
        actual = get_column_value(cell)
        self.assertEqual(expected, actual)

    @parameterized.expand([
        ["B1 gives 1", "B1", 0],
        ["B2 gives 1", "B2", 1],
        ["A1 gives 0", "A1", 0],
        ["Z9 gives 25", "Z9", 8],
        ["F13 gives 12", "F13", 12],
        ["D101 gives 100", "D101", 100]
    ])
    def test_get_row_value(self, name, cell, expected):
        actual = get_row_value(cell)
        self.assertEqual(expected, actual)

    @parameterized.expand([
        ["A1 gives 'PO#'", "A1", "PO#"],
        ["blank gives '[cell not mapped]'", "", "[cell not mapped]"]
    ])
    def test_get_cell_value(self, name, cell, expected):
        actual = get_cell_value(WORKBOOK, cell)
        self.assertEqual(expected, actual)

    def test_get_po_number(self):
        expected = "1234567"
        actual = get_po_number(WORKBOOK, CELLS)
        self.assertEqual(expected, actual)

    def test_get_bill_to_name(self):
        expected = "XYZ Company"
        actual = get_bill_to_name(WORKBOOK, CELLS)
        self.assertEqual(expected, actual)

    def test_get_ship_to_name(self):
        expected = "XYZ Back Door"
        actual = get_ship_to_name(WORKBOOK, CELLS)
        self.assertEqual(expected, actual)

    def test_get_ship_date(self):
        expected = datetime.datetime(2021, 9, 1, 0, 0)  # 09/01/21
        actual = get_ship_date(WORKBOOK, CELLS)
        self.assertEqual(expected, actual)

    def test_get_total_quantity(self):
        expected = 23
        actual = get_total_quantity(WORKBOOK, CELLS)
        self.assertEqual(expected, actual)

    def test_get_total_cost(self):
        expected = 478.00
        actual = get_total_cost(WORKBOOK, CELLS)
        self.assertEqual(expected, actual)


if __name__ == '__main__':
    unittest.main()

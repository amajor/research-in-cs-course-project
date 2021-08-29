import datetime
import unittest
import pandas as pd
from parameterized import parameterized
from spreadsheets.methods import *

FILE_PATH = '../documents/ORDERS.xlsx'
SHEET_NAME = 'Order Sample 1'

"""The workbook to pull values from"""
WORKBOOK = pd.read_excel(
    FILE_PATH,
    sheet_name=SHEET_NAME,
    header=None
)

"""Mapping of the values to cells"""
CELLS = {
    "poNum": "B1",
    "billTo": "B4",
    "shipTo": "B3",
    "shipDate": "B2",
    "totalQty": "B12",
    "totalCost": "G12"
}

"""Fake item table for testing"""
ITEM_TABLE = pd.DataFrame(
    [
        [1, 7, 'CASE', '123', 'Item 1', 20, 140],
        [2, 6, 'PALLET', 124.0, 'Item 2', 18, 108]
    ],
    columns=[
        'LINE_NUM',
        'QTY',
        'UOM',
        'ITEM_NUM',
        'DESC',
        'UNIT_PRICE',
        'COST'
    ]
)


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

    @parameterized.expand([
        ["Row 1 gives line '1'", ITEM_TABLE, 0, 1],
        ["Row 2 gives line '2'", ITEM_TABLE, 1, 2],
        ["Row 3 gives '[not available]'", ITEM_TABLE, 2, '[not available]']
    ])
    def test_get_item_line_number(self, name, table, row, expected):
        actual = get_item_line_number(table, row)
        self.assertEqual(expected, actual)

    @parameterized.expand([
        ["Row 1 gives quantity 7", ITEM_TABLE, 0, 7],
        ["Row 2 gives quantity 6", ITEM_TABLE, 1, 6],
        ["Row 3 gives '[not available]'", ITEM_TABLE, 2, '[not available]']
    ])
    def test_get_item_quantity(self, name, table, row, expected):
        actual = get_item_quantity(table, row)
        self.assertEqual(expected, actual)

    @parameterized.expand([
        ["Row 1 gives 'CASE'", ITEM_TABLE, 0, 'CASE'],
        ["Row 2 gives 'PALLET'", ITEM_TABLE, 1, 'PALLET'],
        ["Row 3 gives '[not available]'", ITEM_TABLE, 2, '[not available]']
    ])
    def test_get_item_unit_of_measure(self, name, table, row, expected):
        actual = get_item_unit_of_measure(table, row)
        self.assertEqual(expected, actual)

    @parameterized.expand([
        ["Row 1 gives item number '123'", ITEM_TABLE, 0, '123'],
        ["Row 2 gives item number '124'", ITEM_TABLE, 1, '124'],
        ["Row 3 gives '[not available]'", ITEM_TABLE, 2, '[not available]']
    ])
    def test_get_item_number(self, name, table, row, expected):
        actual = get_item_number(table, row)
        self.assertEqual(expected, actual)

    @parameterized.expand([
        ["Row 1 gives 'Item 1'", ITEM_TABLE, 0, 'Item 1'],
        ["Row 2 gives 'Item 2'", ITEM_TABLE, 1, 'Item 2'],
        ["Row 3 gives '[not available]'", ITEM_TABLE, 2, '[not available]']
    ])
    def test_get_item_description(self, name, table, row, expected):
        actual = get_item_description(table, row)
        self.assertEqual(expected, actual)

    @parameterized.expand([
        ["Row 1 gives 20.00", ITEM_TABLE, 0, 20.00],
        ["Row 2 gives 18.00", ITEM_TABLE, 1, 18.00],
        ["Row 3 gives '[not available]'", ITEM_TABLE, 2, '[not available]']
    ])
    def test_get_item_unit_price(self, name, table, row, expected):
        actual = get_item_unit_price(table, row)
        self.assertEqual(expected, actual)

    @parameterized.expand([
        ["Row 1 gives 140.00", ITEM_TABLE, 0, 140.00],
        ["Row 2 gives 108.00", ITEM_TABLE, 1, 108.00],
        ["Row 3 gives '[not available]'", ITEM_TABLE, 2, '[not available]']
    ])
    def test_get_item_cost(self, name, table, row, expected):
        actual = get_item_cost(table, row)
        self.assertEqual(expected, actual)


if __name__ == '__main__':
    unittest.main()

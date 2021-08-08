import unittest
from parameterized import parameterized

from spreadsheets.methods import get_column_value, get_row_value


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


if __name__ == '__main__':
    unittest.main()

import pandas as pd

FILE_PATH = '../documents/ORDERS.xlsx'
SHEET_NAME = 'Order Sample 1'

"""The workbook to pull values from"""
WORKBOOK = pd.read_excel(
    FILE_PATH,
    sheet_name=SHEET_NAME,
    header=None
)

"""The table for item information"""
ITEM_TABLE = pd.read_excel(
    FILE_PATH,
    sheet_name=SHEET_NAME,
    usecols="A:G",
    skiprows=5,
    nrows=4
)
"""Mapping column names for item table"""
ITEM_TABLE.columns = [
    'LINE_NUM',    # Line #
    'QTY',         # Quantity
    'UOM',         # Unit of Measure
    'ITEM_NUM',    # Item #
    'DESC',        # Description
    'UNIT_PRICE',  # Unit Price
    'COST'         # Charge
]

"""Mapping of the values to cells"""
ORDER_ONE_CELLS = {
    "poNum": "B1",
    "billTo": "B4",
    "shipTo": "B3",
    "shipDate": "B2",
    "totalQty": "B12",
    "totalCost": "G12"
}

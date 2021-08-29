import pandas as pd

FILE_PATH = '../documents/ORDERS.xlsx'
SHEET_NAME = 'Order Sample 2'

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
    usecols="A:F",
    skiprows=3,
    nrows=4
)
"""Mapping column names for item table"""
ITEM_TABLE.columns = [
    'DESC',        # DESC
    'UNIT_PRICE',  # PRICE
    'QTY',         # QTY
    'UOM',         # U/M
    'ITEM_NUM',    # ITEM_NUM
    'COST'         # COST
]

"""Mapping of the values to cells"""
ORDER_TWO_CELLS = {
    "poNum": "A2",
    "billTo": "",
    "shipTo": "C2",
    "shipDate": "F2",
    "totalQty": "C10",
    "totalCost": "F10"
}

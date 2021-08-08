import pandas as pd

from spreadsheets.methods import get_cell_value, get_po_number, get_ship_date

FILE_PATH = '../documents/ORDERS.xlsx'
SHEET_NAME = 'Order Sample 1'

# The workbook to pull values from
WORKBOOK = pd.read_excel(
    FILE_PATH,
    sheet_name=SHEET_NAME,
    header=None
)

# Mapping of the values to cells
ORDER_ONE_CELLS = {
    "poNum": "B1",
    "billTo": "B4",
    "shipTo": "B3",
    "shipDate": "B2",
    "totalQty": "B12",
    "totalCost": "G12"
}

# - Table of Items
#   - Quantity
#   - Unit of Measure
#   - Item Number
#   - Unit Price
#   - Unit Cost

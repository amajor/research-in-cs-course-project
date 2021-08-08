import pandas as pd

FILE_PATH = '../documents/ORDERS.xlsx'
SHEET_NAME = 'Order Sample 2'

# The workbook to pull values from
WORKBOOK = pd.read_excel(
    FILE_PATH,
    sheet_name=SHEET_NAME,
    header=None
)

# Mapping of the values to cells
ORDER_TWO_CELLS = {
    "poNum": "A2",
    "billTo": "",
    "shipTo": "C2",
    "shipDate": "F2",
    "totalQty": "C10",
    "totalCost": "F10"
}

# - Table of Items
#   - Quantity
#   - Unit of Measure
#   - Item Number
#   - Unit Price
#   - Unit Cost

import string


def get_po_number(workbook, cells):
    """Returns the value for the Purchase Order Number as a string"""
    value = get_cell_value(workbook, cells["poNum"])
    value_as_int = int(value)
    value_as_string = str(value_as_int)
    return value_as_string


def get_bill_to_name(workbook, cells):
    """Returns the value for the Bill-To Name"""
    return get_cell_value(workbook, cells["billTo"])


def get_ship_to_name(workbook, cells):
    """Returns the value for the Ship-To Name (only first line)"""
    value = get_cell_value(workbook, cells["shipTo"])
    value_after_first_line = value.split("\n")
    return value_after_first_line[0]


def get_ship_date(workbook, cells):
    """Returns the value for the Shipping Date"""
    return get_cell_value(workbook, cells["shipDate"])


def get_total_quantity(workbook, cells):
    """Returns the value for the Total Quantity as an integer"""
    value = get_cell_value(workbook, cells["totalQty"])
    value_as_int = int(value)
    return value_as_int


def get_total_cost(workbook, cells):
    """Returns the value for the Total Cost"""
    return get_cell_value(workbook, cells["totalCost"])


def get_cell_value(workbook, cell):
    """Gets the value of the cell (ex: 'B1') from a workbook"""
    if cell == "":
        # No cell was provided.
        return "[cell not mapped]"

    column = get_column_value(cell)
    row = get_row_value(cell)
    return workbook.loc[row].at[column]


def get_column_value(cell):
    """Gets the 'loc' value from a cell"""
    alphabet = dict()
    for index, letter in enumerate(string.ascii_lowercase):
        alphabet[letter] = index

    value = cell[0].lower()
    return alphabet[value]


def get_row_value(cell):
    """Gets the 'at' value from a cell"""
    value = cell[1:]
    return int(value) - 1

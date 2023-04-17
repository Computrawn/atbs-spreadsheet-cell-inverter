#! python3
# spreadsheetCellInverter.py — An exercise in manipulating Excel files.
# For more information, see project_details.txt.

import openpyxl
from openpyxl.utils import get_column_letter

file_name = input("Please enter filename here: ")
xy_list = []


def cell_inverter(document):
    """Opens xlsx file, inverts its cells and saves the inverted file."""
    wb = openpyxl.load_workbook(f"{document}.xlsx")
    sheet = wb.active

    for column in range(1, sheet.max_column + 1):
        column_letter = get_column_letter(column)
        column_list = []
        for row in range(1, sheet.max_row + 1):
            column_list.append(sheet[f"{column_letter}{row}"].value)
        xy_list.append(column_list)

    wb.save(f"{document}_inverted.xlsx")


cell_inverter(file_name)
print(xy_list)

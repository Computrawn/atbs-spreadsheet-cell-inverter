#! python3
# spreadsheetCellInverter.py — An exercise in manipulating Excel files.
# For more information, see project_details.txt.

import openpyxl

file_name = input("Please enter filename here: ")


def cell_inverter(document):
    wb = openpyxl.load_workbook(f"{document}.xlsx")
    sheet = wb.active
    # TODO: write nested loops to invert document cells.
    wb.save(f"{document}_inverted.xlsx")


cell_inverter(file_name)

#! python3
# spreadsheetCellInverter.py — An exercise in manipulating Excel files.
# For more information, see project_details.txt.

import openpyxl

file_name = input('Please enter filename here: ')

wb = openpyxl.load_workbook(file_name)
sheet = wb.active

wb.save(file_name)
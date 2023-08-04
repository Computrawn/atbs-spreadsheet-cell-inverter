
# Spreadsheet Cell Inverter

Write a program to invert the row and column of the cells in the spreadsheet. 

For example, the value at row 5, column 3 will be at row 3, column 5 (and vice versa). 

This should be done for all cells in the spreadsheet. 

For example, the “before” and “after” spreadsheets would look something like Figure 13-13.

You can write this program by using nested for loops to read the spreadsheet’s data into a list of lists data structure. 

This data structure could have sheetData\[x][y] for the cell at column x and row y. 

Then, when writing out the new spreadsheet, use sheetData\[y][x] for the cell at column x and row y.

**Excerpt From Automate the Boring Stuff with Python: Practical Programming for Total Beginners, 2nd Edition  
Al Sweigart
This material may be protected by copyright.**
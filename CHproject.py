"""Created for use at Carlyle House, 2020.
Can be edited for use at any museum institution to help sort files for researchers.  """

#read an excel file
from openpyxl import Workbook

# Program to extract number 
# of rows using Python   
# Give the location of the file 
loc = input("Please enter the name of your file: ")
  
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0)
# For row 0 and column 0 
sheet.cell_value(0, 0) 
  
# Extracting number of rows 
print(sheet.nrows)

# Extracting number of columns 
print(sheet.ncols)

# Program extracting all columns 
# name in Python 
for i in range(sheet.ncols): 
    print(sheet.cell_value(0, i))
    
    
# Program extracting first column
for i in range(sheet.nrows): 
    print(sheet.cell_value(i, 0))
   
# Program to extract a particular row value  
print(sheet.row_values(1)) 
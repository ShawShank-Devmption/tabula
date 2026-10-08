#!/usr/bin/env python3
from openpyxl import load_workbook
from copy import copy

# Load the workbook
wb = load_workbook('/Users/shashankbhat/Projects/compiler_project/demo/runs/live-haiku/t2_insert_row__default__1/book.xlsx')

sales_sheet = wb['Sales']
summary_sheet = wb['Summary']

# Insert a new row at position 11 (pushes current row 11 and below down)
sales_sheet.insert_rows(11, 1)

# Add the new sale data to row 11
sales_sheet['A11'] = 'West'
sales_sheet['B11'] = 'Dana Lee'
sales_sheet['C11'] = 40
sales_sheet['D11'] = 125

# Add formulas for Revenue and Commission (matching the pattern from other rows)
sales_sheet['E11'] = '=C11*D11'
sales_sheet['F11'] = '=E11*CommissionRate'

# Now the totals row has moved from row 22 to row 23
# Update the SUM formulas in row 23 to include the new row
sales_sheet['C23'] = '=SUM(C2:C22)'
sales_sheet['E23'] = '=SUM(E2:E22)'
sales_sheet['F23'] = '=SUM(F2:F22)'

# Update the named range SalesRevenue to include the new row
# Change from Sales!$E$2:$E$21 to Sales!$E$2:$E$22
wb.defined_names['SalesRevenue'].value = 'Sales!$E$2:$E$22'

# Update Summary sheet formulas that reference row 22 (now row 23)
summary_sheet['B3'] = '=Sales!E23'  # Total revenue
summary_sheet['B4'] = '=Sales!F23'  # Total commission

# The SUMIF for North revenue needs to be updated to include row 22 in the range
summary_sheet['B6'] = '=SUMIF(Sales!A2:A22,"North",Sales!E2:E22)'

# Save the workbook
wb.save('/Users/shashankbhat/Projects/compiler_project/demo/runs/live-haiku/t2_insert_row__default__1/book.xlsx')

print("✓ Inserted new sale row at row 11")
print("✓ Added data: West, Dana Lee, 40 units, $125 unit price")
print("✓ Formulas for Revenue and Commission added")
print("✓ Updated total row formulas (now row 23)")
print("✓ Updated Summary sheet references")
print("✓ Updated SalesRevenue named range")
print("✓ Saved book.xlsx")

wb.close()

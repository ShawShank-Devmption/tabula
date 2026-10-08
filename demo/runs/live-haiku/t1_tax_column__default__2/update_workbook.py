#!/usr/bin/env python3
import openpyxl

# Load the workbook
wb = openpyxl.load_workbook('book.xlsx')

# Get the Sales sheet
sales_ws = wb['Sales']

# Add header "Tax" in G1
sales_ws['G1'] = 'Tax'

# Add tax formulas for rows 2-21
# Tax = Revenue (column E) * Tax rate (from Inputs sheet B2)
for row in range(2, 22):
    sales_ws[f'G{row}'] = f'=E{row}*Inputs!$B$2'

# Add total formula in G22
sales_ws['G22'] = '=SUM(G2:G21)'

# Get the Summary sheet
summary_ws = wb['Summary']

# Add "Total tax" label in A8
summary_ws['A8'] = 'Total tax'

# Add formula in B8 that references the total tax from Sales sheet
summary_ws['B8'] = '=Sales!G22'

# Save the workbook
wb.save('book.xlsx')

print("✓ Added Tax column (G) to Sales sheet with formulas")
print("✓ Added Tax total formula in G22")
print("✓ Added Total tax label and formula to Summary sheet row 8")
print("✓ Saved changes to book.xlsx")

#!/usr/bin/env python3
from openpyxl import load_workbook

# Load the workbook
wb = load_workbook('book.xlsx')
sales_sheet = wb['Sales']
summary_sheet = wb['Summary']

# Add "Tax" header in G1
sales_sheet['G1'] = 'Tax'

# Add Tax formulas for each sales row (rows 2-21)
# Tax = Revenue (column E) * TaxRate
for row in range(2, 22):
    sales_sheet[f'G{row}'] = f'=E{row}*TaxRate'

# Add total tax formula in G22 (the Total row)
sales_sheet['G22'] = '=SUM(G2:G21)'

# Add "Total tax" label in A8 of Summary sheet
summary_sheet['A8'] = 'Total tax'

# Add formula in B8 to reference the total tax from Sales sheet
summary_sheet['B8'] = '=Sales!G22'

# Save the workbook
wb.save('book.xlsx')
print("Workbook updated successfully!")
print("- Added 'Tax' column (G) to Sales sheet with formulas")
print("- Added 'Total tax' row (row 8) to Summary sheet")

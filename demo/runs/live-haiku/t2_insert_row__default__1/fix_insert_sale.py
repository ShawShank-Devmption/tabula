#!/usr/bin/env python3
from openpyxl import load_workbook
import re

# Reload original file
wb = load_workbook('/Users/shashankbhat/Projects/compiler_project/demo/runs/live-haiku/t2_insert_row__default__1/book.xlsx')

# Close and reload to get clean state
wb.close()

# Reload to start fresh
wb = load_workbook('/Users/shashankbhat/Projects/compiler_project/demo/runs/live-haiku/t2_insert_row__default__1/book.xlsx')
sales_sheet = wb['Sales']
summary_sheet = wb['Summary']

# Manual approach: Copy rows down from bottom to top to avoid conflicts
# First, let's copy all data from rows 11-22 to rows 12-23
for col_idx in range(1, 7):  # columns A-F
    for row_idx in range(22, 10, -1):  # Go from 22 down to 11
        source_cell = sales_sheet.cell(row_idx, col_idx)
        dest_cell = sales_sheet.cell(row_idx + 1, col_idx)

        # Copy value and formula
        if source_cell.value is not None:
            if isinstance(source_cell.value, str) and source_cell.value.startswith('='):
                # It's a formula - increment row numbers in it
                formula = source_cell.value
                # Replace references like C11 with C12, but be careful with ranges like C2:C21
                def increment_row(match):
                    cell_ref = match.group(0)
                    # Extract column letter and row number
                    col_match = re.match(r'([A-Z]+)(\d+)', cell_ref)
                    if col_match:
                        col = col_match.group(1)
                        row = int(col_match.group(2))
                        return f"{col}{row + 1}"
                    return cell_ref

                # Replace all cell references with incremented row numbers
                new_formula = re.sub(r'[A-Z]+\d+', increment_row, formula)
                dest_cell.value = new_formula
            else:
                dest_cell.value = source_cell.value

# Now set row 11 with new data
sales_sheet['A11'] = 'West'
sales_sheet['B11'] = 'Dana Lee'
sales_sheet['C11'] = 40
sales_sheet['D11'] = 125
sales_sheet['E11'] = '=C11*D11'
sales_sheet['F11'] = '=E11*CommissionRate'

# Clear row 23 first, then set totals
for col_idx in range(1, 7):
    sales_sheet.cell(23, col_idx).value = None

# Set totals in row 23
sales_sheet['A23'] = 'Total'
sales_sheet['C23'] = '=SUM(C2:C22)'
sales_sheet['E23'] = '=SUM(E2:E22)'
sales_sheet['F23'] = '=SUM(F2:F22)'

# Update the named range SalesRevenue
wb.defined_names['SalesRevenue'].value = 'Sales!$E$2:$E$22'

# Update Summary sheet formulas
summary_sheet['B3'] = '=Sales!E23'  # Total revenue
summary_sheet['B4'] = '=Sales!F23'  # Total commission
summary_sheet['B6'] = '=SUMIF(Sales!A2:A22,"North",Sales!E2:E22)'  # North revenue

# Save the workbook
wb.save('/Users/shashankbhat/Projects/compiler_project/demo/runs/live-haiku/t2_insert_row__default__1/book.xlsx')

print("✓ Fixed: Inserted new sale row at row 11 with proper formula adjustments")
print("✓ All formulas in shifted rows are now correctly updated")
print("✓ Saved book.xlsx")

wb.close()

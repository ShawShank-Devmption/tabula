#!/usr/bin/env python3
from openpyxl import load_workbook

# Load the workbook
wb = load_workbook('/Users/shashankbhat/Projects/compiler_project/demo/runs/live-haiku/t2_insert_row__default__1/book.xlsx')

# List all sheets
print("Sheets in workbook:", wb.sheetnames)

# Examine the Sales sheet
sales_sheet = wb['Sales']
print("\n=== Sales Sheet ===")
print("Max row:", sales_sheet.max_row)
print("Max column:", sales_sheet.max_column)

# Print all data with formulas
print("\nContent (first 15 rows):")
for row_idx in range(1, min(16, sales_sheet.max_row + 1)):
    row_data = []
    for col_idx in range(1, sales_sheet.max_column + 1):
        cell = sales_sheet.cell(row_idx, col_idx)
        # Show both value and formula if it's a formula
        if cell.value is not None:
            if isinstance(cell.value, str) and cell.value.startswith('='):
                row_data.append(f"[{cell.value}]")
            else:
                row_data.append(str(cell.value))
        else:
            row_data.append("")
    print(f"Row {row_idx}: {row_data}")

# Check if there's a Summary sheet
if 'Summary' in wb.sheetnames:
    summary_sheet = wb['Summary']
    print("\n=== Summary Sheet ===")
    print("Max row:", summary_sheet.max_row)
    print("Max column:", summary_sheet.max_column)
    print("\nContent:")
    for row_idx in range(1, summary_sheet.max_row + 1):
        row_data = []
        for col_idx in range(1, summary_sheet.max_column + 1):
            cell = summary_sheet.cell(row_idx, col_idx)
            if cell.value is not None:
                if isinstance(cell.value, str) and cell.value.startswith('='):
                    row_data.append(f"[{cell.value}]")
                else:
                    row_data.append(str(cell.value))
            else:
                row_data.append("")
        print(f"Row {row_idx}: {row_data}")

wb.close()

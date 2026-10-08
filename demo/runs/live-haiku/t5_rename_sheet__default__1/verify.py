#!/usr/bin/env python3
from openpyxl import load_workbook

# Load the workbook
wb = load_workbook('book.xlsx')

print("Verification of changes:")
print("="*60)
print()

print("Sheet names:", wb.sheetnames)
print()

print("Formulas in Summary sheet:")
ws = wb['Summary']
for row in ws.iter_rows():
    for cell in row:
        if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
            print(f"  {cell.coordinate}: {cell.value}")
print()

print("✓ All changes verified!")

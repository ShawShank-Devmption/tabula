#!/usr/bin/env python3
from openpyxl import load_workbook

# Load the workbook
wb = load_workbook('book.xlsx')
ws_summary = wb['Summary']

print("Before fixes:")
print(f"B3 (Total revenue): {ws_summary['B3'].value}")
print(f"B4 (Total commission): {ws_summary['B4'].value}")

# Fix B3: Total revenue should sum all revenues from Sales sheet
ws_summary['B3'].value = '=SUM(Sales!E2:E21)'

# Fix B4: Total commission should sum all commissions from Sales sheet
ws_summary['B4'].value = '=SUM(Sales!F2:F21)'

print("\nAfter fixes:")
print(f"B3 (Total revenue): {ws_summary['B3'].value}")
print(f"B4 (Total commission): {ws_summary['B4'].value}")

# Save the workbook
wb.save('book.xlsx')
print("\nWorkbook saved!")

wb.close()

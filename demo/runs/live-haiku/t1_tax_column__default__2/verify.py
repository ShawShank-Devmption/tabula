#!/usr/bin/env python3
import openpyxl

# Load the updated workbook
wb = openpyxl.load_workbook('book.xlsx')

print("=== Sales Sheet ===")
sales_ws = wb['Sales']
# Show header and first few rows
print(f"G1 (header): {sales_ws['G1'].value}")
print(f"G2 (formula): {sales_ws['G2'].value}")
print(f"G22 (total): {sales_ws['G22'].value}")

print("\n=== Summary Sheet ===")
summary_ws = wb['Summary']
print(f"A8: {summary_ws['A8'].value}")
print(f"B8: {summary_ws['B8'].value}")

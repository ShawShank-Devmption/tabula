from openpyxl import load_workbook

# Load the workbook
wb = load_workbook('book.xlsx')

# Get all named ranges
print("Named ranges:")
print(f"Defined names: {wb.defined_names}")

# Try to iterate
for name in wb.defined_names:
    print(f"Name: {name}")

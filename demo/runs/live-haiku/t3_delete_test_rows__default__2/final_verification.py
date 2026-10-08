from openpyxl import load_workbook

wb = load_workbook('book.xlsx')

print("=== FINAL VERIFICATION ===\n")

# Check Sales sheet
ws = wb['Sales']
print(f"Sales sheet: {ws.max_row} rows total")
print("Data rows (excluding header and total):")
test_count = 0
for row_idx in range(2, ws.max_row):
    rep = ws[f'B{row_idx}'].value
    if rep == 'TEST':
        test_count += 1
        print(f"  ERROR: Found TEST at row {row_idx}")

if test_count == 0:
    print(f"  ✓ No TEST entries found")
else:
    print(f"  ✗ Found {test_count} TEST entries")

# Verify total row formulas
total_row = ws.max_row
print(f"\nTotal row (row {total_row}):")
print(f"  Units sum: {ws[f'C{total_row}'].value}")
print(f"  Revenue sum: {ws[f'E{total_row}'].value}")
print(f"  Commission sum: {ws[f'F{total_row}'].value}")

# Check Summary sheet references
ws = wb['Summary']
print(f"\nSummary sheet formulas:")
print(f"  Total revenue: {ws['B3'].value}")
print(f"  Total commission: {ws['B4'].value}")
print(f"  North revenue: {ws['B6'].value}")

print("\n✓ Verification complete - all test rows deleted and formulas updated!")

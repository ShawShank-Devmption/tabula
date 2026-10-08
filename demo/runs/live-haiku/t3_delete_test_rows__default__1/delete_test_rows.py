from openpyxl import load_workbook

# Load the workbook
wb = load_workbook('book.xlsx')
sales_ws = wb['Sales']
summary_ws = wb['Summary']

# Delete rows 15 and 8 (in reverse order to avoid row shifting)
sales_ws.delete_rows(15, 1)
sales_ws.delete_rows(8, 1)

# Update the named range SalesRevenue from E2:E21 to E2:E19
for defined_name in wb.defined_names:
    if defined_name == 'SalesRevenue':
        wb.defined_names['SalesRevenue'].attr_text = 'Sales!$E$2:$E$19'

# Update formulas in Summary sheet
# Row 3: Total revenue formula - change E22 to E20
summary_ws['B3'].value = '=Sales!E20'

# Row 4: Total commission formula - change F22 to F20
summary_ws['B4'].value = '=Sales!F20'

# Row 6: North revenue formula - change A2:A21 and E2:E21 to A2:A19 and E2:E19
summary_ws['B6'].value = '=SUMIF(Sales!A2:A19,"North",Sales!E2:E19)'

# Also update the total row formulas in Sales sheet (row 20 is now the total after deletion)
# C20: =SUM(C2:C19)
sales_ws['C20'].value = '=SUM(C2:C19)'
# E20: =SUM(E2:E19)
sales_ws['E20'].value = '=SUM(E2:E19)'
# F20: =SUM(F2:F19)
sales_ws['F20'].value = '=SUM(F2:F19)'

# Save the workbook
wb.save('book.xlsx')

print("Test rows deleted successfully!")
print("Updated formulas and named ranges.")

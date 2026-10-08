"""T5 with plain openpyxl: the straightforward approach (scripted reference, naive)."""
import openpyxl

wb = openpyxl.load_workbook("book.xlsx")
wb["Sales"].title = "Q3 Sales"
wb.save("book.xlsx")
print("renamed Sales to Q3 Sales")

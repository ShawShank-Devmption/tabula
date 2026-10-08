"""T4 with plain openpyxl (scripted reference for the default arm; correct)."""
import openpyxl

wb = openpyxl.load_workbook("book.xlsx")
s = wb["Summary"]
s["B4"] = "=Sales!F22"
s["B6"] = '=SUMIF(Sales!A2:A21,"North",Sales!E2:E21)'
wb.save("book.xlsx")
print("fixed Summary!B4 and B6")

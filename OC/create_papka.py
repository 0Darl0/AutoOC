import openpyxl

wb = openpyxl.Workbook()
page = wb.active
page.append("0/13")
wb.save("1.xlsx")
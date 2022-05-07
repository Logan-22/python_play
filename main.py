from openpyxl import Workbook, load_workbook

wb = load_workbook('excel/test_python.xlsx')
ws = wb.active
ws['G7'].value = "Loganand"

wb.save('excel/test_python.xlsx')

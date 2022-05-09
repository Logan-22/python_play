from openpyxl import Workbook, load_workbook

# To create New sheet
# wb = load_workbook('excel/test_python.xlsx')
# ws = wb['Marks']

# wb.create_sheet("Test")
# wb.save('excel/test_python.xlsx')
# print(wb.sheetnames)

# To create New Workbook

wb = Workbook()
ws = wb.active
ws.title = "Data"

ws.append(["Loganand", "Is", "Great", "!"])

wb.save("excel/Test.xlsx")

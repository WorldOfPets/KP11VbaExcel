import openpyxl

wb = openpyxl.load_workbook("testxl.xlsm", read_only=False, keep_vba=True)
print(wb.sheetnames)
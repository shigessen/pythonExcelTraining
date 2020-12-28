import openpyxl

#wb = openpyxl.load_workbook(r"d:\python\training\data\sample.xlsx")
wb = openpyxl.load_workbook(r".\data\sample.xlsx")
for sheet in wb:
    for row in sheet:
        for cell in row:
            print(cell.value)

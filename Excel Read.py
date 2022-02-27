import xlrd

# Author : Imtiaz Adar
# Project : Excel Using Python
# Language : Python

path = "C:\\Users\\imtia\\PycharmProjects\\Excel\\venv\\informations_.xlsx"
work_book = xlrd.open_workbook(path)
sheet = work_book.sheet_by_index(0)

print()
for row in range(sheet.nrows):
    for col in range(sheet.ncols):
        if col < sheet.ncols - 1:
            print(sheet.cell_value(row, col), end=' | ')
        else:
            print(sheet.cell_value(row, col),end='')
    print()

from numpy import var
import openpyxl
import pprint

read_excel = './file.xlsx'       #khai báo file làm CO
wb = openpyxl.load_workbook(read_excel)
sheet = wb['Sheet8']            #Tên sheet file CO
print('Nhap so dong hang')
donghang = int(input())
g = sheet.iter_rows(min_row=1, max_row=13, min_col=1, max_col=8)
print(type(g))
# <class 'generator'>
cells_list=list(g)
i = 1
j = 1
print("dòng đầu là")
print(cells_list[1][1].value)
for i in range(donghang):
    for j in range(6):
        print(cells_list[i+1][j+1].value)

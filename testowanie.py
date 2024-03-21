import openpyxl
import sys
sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('DaneTestoweCHATGPT.xlsx')
sheet = wb.active

first_row = sheet.iter_rows(min_row=1, max_row=1)

for cell_tuple in first_row:
    for cell in cell_tuple:
        print(cell.value)



# for cell in first_row:
#   print(cell.value)
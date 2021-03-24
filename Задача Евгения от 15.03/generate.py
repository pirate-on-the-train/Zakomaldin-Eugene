from openpyxl import Workbook
import random

wb = Workbook()
ws = wb.active

for i in range(1, 11):
    for j in range(1, 11):
        ws.cell(column = j, row = i, value = random.uniform(0.1, 100.0))

wb.save('1.xlsx')
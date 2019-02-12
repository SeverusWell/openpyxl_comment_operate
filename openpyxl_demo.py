# -*- coding: utf-8 -*-

from openpyxl import load_workbook
from openpyxl import Workbook

source_wb = load_workbook("input.xlsx")

source_sheet = source_wb["Sheet1"]
# source_sheet = source_wb.get_sheet_by_name("Sheet1")

target_wb = Workbook()
target_sheet = target_wb.active

for row in source_sheet.iter_rows():
    for cell in row:
        print(cell.value)
        # 内容
        target_sheet[cell.column + str(cell.row * 2 - 1)] = str(cell.value) if cell.value else ''
        # 备注
        target_sheet[cell.column + str(cell.row * 2)] = str(cell.comment) if cell.comment else ''

target_wb.save('output.xlsx')

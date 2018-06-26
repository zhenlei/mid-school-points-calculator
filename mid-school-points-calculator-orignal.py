from mmap import mmap, ACCESS_READ
from xlrd import open_workbook, cellname, XL_CELL_TEXT
import numpy as np
from xlsxwriter import Workbook

titles = []
factors = [1, 0.8, 0.6]

wb_writer = Workbook('result.xlsx')
sheet_writer = wb_writer.add_worksheet("results")

wb = open_workbook('class.xlsx')
sheet = wb.sheet_by_name('class')

for row_index in range(0, sheet.nrows):
    cell = sheet.cell(row_index, 0)
    if (cell.ctype != XL_CELL_TEXT):
        break

    points = sheet.row_values(row_index, 0, 8)
    if (row_index == 0):
        sheet_writer.write_row(row_index, 0, points)
        continue

    sub_points = points[5:]
    sub_points.sort(reverse=True)

    sub_points = sub_points * np.asarray(factors)

    post_points = points[:5] + sub_points.tolist()
    print("post points: ", post_points)

    points.append(sum(post_points[1:]))
    print("points: ", points)
    sheet_writer.write_row(row_index, 0, points)

wb_writer.close()

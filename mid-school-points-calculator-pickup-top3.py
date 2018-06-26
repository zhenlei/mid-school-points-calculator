from mmap import mmap, ACCESS_READ
from xlrd import open_workbook, cellname, XL_CELL_TEXT, XL_CELL_EMPTY, XL_CELL_BLANK
import numpy as np
from xlsxwriter import Workbook

row_start = 3
row_end = 13
sub_start = 5
titles = []
factors = [1, 0.8, 0.6]

wb_writer = Workbook('result.xlsx')
sheet_writer = wb_writer.add_worksheet("results")

wb = open_workbook('class.xlsx')
sheet = wb.sheet_by_index(0)

for row_index in range(0, sheet.nrows):
    cell = sheet.cell(row_index, 0)
    # if (cell.ctype != XL_CELL_TEXT):
    if (cell.ctype == XL_CELL_EMPTY or cell.ctype == XL_CELL_BLANK):
        print("--------------------break on", row_index)
        print("the cell type is " , cell.ctype)
        break

    print("the cell type is " , cell.ctype)
    points = sheet.row_values(row_index, row_start, row_end)
    if (row_index == 0):
        sheet_writer.write_row(row_index, 0, [points[0], "summary"])
        continue

    print("points is ", points)
    sub_points = points[sub_start:]
    sub_points.sort(reverse=True)
    sub_points = sub_points[:3]
    print(sub_points)

    sub_points = sub_points * np.asarray(factors)

    post_points = points[:sub_start] + sub_points.tolist()
    print("post points: ", post_points)

    # points.append(sum(post_points[2:]))
    present_points = [points[0], sum(post_points[2:])]
    print("points: ", present_points)
    sheet_writer.write_row(row_index, 0, present_points)

wb_writer.close()

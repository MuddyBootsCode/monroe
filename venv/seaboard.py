import os
import openpyxl
from openpyxl import load_workbook
import csv
from openpyxl import Workbook
import datetime
import operator
from operator import itemgetter
from decimal import *

os.chdir('/Users/porte/Desktop')

wb = load_workbook(filename='seaboard.xlsx')
ws = wb.active

headings = ["REF", "CUSTOMER:JOB", "SHIP DATE", "DUE DATE", "ITEM", "GROSS VOL", "GROSS EXP", "GROSS OWNER"]

lease_name_tags = []

for cell in ws["A"]:
    try:
        if "County" in ws['A' + str(cell.row + 1)].value:
            lease_name_tags.append(cell)
    except TypeError:
        continue

for cell in ws['A']:
    try:
        if "Grand Total Share - Oil" in cell.value:
            lease_name_tags.append(cell)
    except TypeError:
        continue

sheet_date = "4/30/2018"
sheet_date = datetime.datetime.strptime(sheet_date, '%m/%d/%Y')
wb.create_sheet("Totals")
active_ws = wb['Totals']
for idx, title in enumerate(headings):
    active_ws.cell(row=1, column=idx + 1, value=title)

row_counter = 2

for idx, cell in enumerate(lease_name_tags):
    print(idx, len(lease_name_tags))
    try:
        for x in range(lease_name_tags[idx].row, lease_name_tags[idx + 1].row - 1):
            if "Share Total" in ws['A' + str(x)].value:
                active_ws.cell(row=row_counter, column=2, value="APC:" + lease_name_tags[idx].value.split("-")[0])
                active_ws.cell(row=row_counter + 1, column=2, value="APC:" + lease_name_tags[idx].value.split("-")[0])
                active_ws.cell(row=row_counter, column=3, value=sheet_date)
                active_ws.cell(row=row_counter, column=4, value=sheet_date)
                if ws['B' + str(x)].value == "Oil":
                    active_ws.cell(row=row_counter, column=5, value="Texas Oil")
                    active_ws.cell(row=row_counter + 1, column=5, value="Texas Severence Exp")
                else:
                    active_ws.cell(row=row_counter, column=5, value="TEXAS Gas")
                    active_ws.cell(row=row_counter + 1, column=5, value="Texas Severence Exp")

                row_counter += 2
    except IndexError:
        if "Share Total" in ws['A' + str(x)].value:
            active_ws.cell(row=row_counter, column=2, value="APC:" + lease_name_tags[idx].value.split("-")[0])
            active_ws.cell(row=row_counter + 1, column=2, value="APC:" + lease_name_tags[idx].value.split("-")[0])
            active_ws.cell(row=row_counter, column=3, value=sheet_date)
            active_ws.cell(row=row_counter, column=4, value=sheet_date)
            if ws['B' + str(x)].value == "Oil":
                active_ws.cell(row=row_counter, column=5, value="Texas Oil")
                active_ws.cell(row=row_counter + 1, column=5, value="Texas Severence Exp")
            else:
                active_ws.cell(row=row_counter, column=5, value="TEXAS Gas")
                active_ws.cell(row=row_counter + 1, column=5, value="Texas Severence Exp")

            row_counter += 2

# print([cell.value for cell in lease_name_tags])

wb.save('test.xlsx')
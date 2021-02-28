import pandas as pd
import xlrd 
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
math = xlrd.open_workbook("insert_sheet_name.xlsx")
math_sheets = []
for x in range(math.nsheets):
    math_sheets.append(math.sheet_by_index(x))
print("The num of sheeets is ", len(math_sheets))
tutors = xlrd.open_workbook("insert_sheet_name.xlsx")
names = []
for x in range(2, tutors.sheet_by_index(0).nrows):
    names.append(tutors.sheet_by_index(0).row_values(x)[0])
    name = names[x-2].split(",")
    name = name[1] + " " + name[0]
    name = str(name).strip()
    if name == "wrong name":
        name="fixed name"

    names[x-2] = name.strip()

tutor_attendance = {}
for name in names:
    tutor_attendance[name] = [0, ""]
for n in names:
    for x in math_sheets:
        for rowidx in range(x.nrows):
            row = x.row(rowidx)
            for colidx, cell in enumerate(row):
                if str(cell.value).strip() == n:
                    tutor_attendance[n][1] = "Math"
                if str(cell.value).strip() == n and x.cell_value(rowx = rowidx, colx = colidx+1)==True:
                    tutor_attendance[n][0]+=1

math = xlrd.open_workbook("insert_sheet_name.xlsx")
math_sheets = []
for x in range(math.nsheets):
    math_sheets.append(math.sheet_by_index(x))
print("The num of sheeets is ", len(math_sheets))

for n in names:
    for x in math_sheets:
        for rowidx in range(x.nrows):
            row = x.row(rowidx)
            for colidx, cell in enumerate(row):
                if str(cell.value).strip() == n:
                    tutor_attendance[n][1] = "Science"
                if str(cell.value).strip() == n and x.cell_value(rowx = rowidx, colx = colidx+1)==True:
                    tutor_attendance[n][0]+=1

math = xlrd.open_workbook("insert_sheet_name.xlsx")
math_sheets = []
for x in range(math.nsheets):
    math_sheets.append(math.sheet_by_index(x))
print("The num of sheeets is ", len(math_sheets))

for n in names:
    for x in math_sheets:
        for rowidx in range(x.nrows):
            row = x.row(rowidx)
            for colidx, cell in enumerate(row):
                if str(cell.value).strip() == n:
                    tutor_attendance[n][1] = "SS"
                if str(cell.value).strip() == n and x.cell_value(rowx = rowidx, colx = colidx+1)==True:
                    tutor_attendance[n][0]+=1


math = xlrd.open_workbook("insert_sheet_name.xlsx")
math_sheets = []
for x in range(math.nsheets):
    math_sheets.append(math.sheet_by_index(x))
print("The num of sheeets is ", len(math_sheets))

for n in names:
    for x in math_sheets:
        for rowidx in range(x.nrows):
            row = x.row(rowidx)
            for colidx, cell in enumerate(row):
                if str(cell.value).strip() == n:
                    tutor_attendance[n][1] = "CA"
                if str(cell.value).strip() == n and x.cell_value(rowx = rowidx, colx = colidx+1)==True:
                    tutor_attendance[n][0]+=1


math = xlrd.open_workbook("insert_sheet_name.xlsx")
math_sheets = []
for x in range(math.nsheets):
    math_sheets.append(math.sheet_by_index(x))
print("The num of sheeets is ", len(math_sheets))

for n in names:
    for x in math_sheets:
        for rowidx in range(x.nrows):
            row = x.row(rowidx)
            for colidx, cell in enumerate(row):
                if str(cell.value).strip() == n:
                    tutor_attendance[n][1] = "WL"
                if str(cell.value).strip() == n and x.cell_value(rowx = rowidx, colx = colidx+1)==True:
                    tutor_attendance[n][0]+=1

math = xlrd.open_workbook("insert_sheet_name.xlsx")
math_sheets = []
for x in range(math.nsheets):
    math_sheets.append(math.sheet_by_index(x))
print("The num of sheeets is ", len(math_sheets))

for n in names:
    for x in math_sheets:
        for rowidx in range(x.nrows):
            row = x.row(rowidx)
            for colidx, cell in enumerate(row):
                if str(cell.value).strip() == n:
                    tutor_attendance[n][1] = "Guided Study"
                if str(cell.value).strip() == n and x.cell_value(rowx = rowidx, colx = colidx+1)==True:
                    tutor_attendance[n][0]+=1

math = xlrd.open_workbook("insert_sheet_name.xlsx")
math_sheets = []
for x in range(math.nsheets):
    math_sheets.append(math.sheet_by_index(x))
print("The num of sheeets is ", len(math_sheets))

for n in names:
    for x in math_sheets:
        for rowidx in range(x.nrows):
            row = x.row(rowidx)
            for colidx, cell in enumerate(row):
                if str(cell.value).strip() == n:
                    tutor_attendance[n][1] = "CA"
                if str(cell.value).strip() == n and x.cell_value(rowx = rowidx, colx = colidx+1)==True:
                    tutor_attendance[n][0]+=1

math = xlrd.open_workbook("insert_sheet_name.xlsx")
math_sheets = []
for x in range(math.nsheets):
    math_sheets.append(math.sheet_by_index(x))
print("The num of sheeets is ", len(math_sheets))

for n in names:
    for x in math_sheets:
        for rowidx in range(x.nrows):
            row = x.row(rowidx)
            for colidx, cell in enumerate(row):
                if str(cell.value).strip() == n:
                    tutor_attendance[n][1] = "DW"
                if str(cell.value).strip() == n and x.cell_value(rowx = rowidx, colx = colidx+1)==True:
                    tutor_attendance[n][0]+=1

math = xlrd.open_workbook("insert_sheet_name.xlsx")
math_sheets = []
for x in range(math.nsheets):
    math_sheets.append(math.sheet_by_index(x))
print("The num of sheeets is ", len(math_sheets))

for n in names:
    for x in math_sheets:
        for rowidx in range(x.nrows):
            row = x.row(rowidx)
            for colidx, cell in enumerate(row):
                if str(cell.value).strip() == n:
                    tutor_attendance[n][1] = "CS"
                if str(cell.value).strip() == n and x.cell_value(rowx = rowidx, colx = colidx+1)==True:
                    tutor_attendance[n][0]+=1

print(tutor_attendance)
print()
print()
for x in tutor_attendance.values():
    print(x[1])

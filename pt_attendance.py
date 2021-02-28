import pandas as pd
import xlrd 
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
math = xlrd.open_workbook("/Users/krishna/Documents/pt_attendance/FNL-AT-AT PT MATH.xlsx")
math_sheets = []
for x in range(math.nsheets):
    math_sheets.append(math.sheet_by_index(x))
print("The num of sheeets is ", len(math_sheets))
tutors = xlrd.open_workbook("/Users/krishna/Documents/pt_attendance/Peer Tutors Roster Sem1 20-21.xlsx")
names = []
for x in range(2, tutors.sheet_by_index(0).nrows):
    names.append(tutors.sheet_by_index(0).row_values(x)[0])
    name = names[x-2].split(",")
    name = name[1] + " " + name[0]
    name = str(name).strip()
    if name == "Ribhav Bhatia":
        name="Ribhiv Bhatia"
    if name == "Aniketh Bhaskar":
        name = "Aniketh Bhaska"
    if name == "Seunghee Han":
        name = "Olivia Han"
    if name == "Dylan Heydenburg":
        name = "Dylan Hydensburg"
    if name == "Sahil Kamat":
        name = "Sahit Kamat"
    if name == "Gyeongwon Kim":
        name = "Gyeongwon (James) Kim"
    if name == "Seohyun Kim":
        name = "Bella Kim"
    if name == "Ha Eun Kwon":
        name = "Esther Kwon"
    if name == "Regina Lapitsky":
        name = "Rehina (Regina) Lapytska"
    if name == "Haejoon Lim":
        name = "John Lim"
    if name == "Jonathan Mann":
        name = "Jonathan (Yoni) Mann"
    if name == "Caroline Mazur Sarocka":
        name = "Caroline Mazur-Sarocka"
    if name == "Abigail Minin":
        name = "Abby Minin"
    if name == "Madeline Mitchell":
        name = "Maddie Mitchell"
    if name == "Varsha Mullangi":
        name = "Varsha Mulangi"
    if name == "Aditya Nair":
        name = "Adi Nair"
    if name == "Sudha Nallacheruvu":
        name = "Sailaja Nallacheruvu"
    if name == "Joshua Neela":
        name = "Josh Neela"
    if name == "Alexandra Sokolowski":
        name = "Alexandra Sokolowki"
    if name == "Swetha Subramanian":
        name = "Swetha Subramanium"
    if name == "Samuel Sweet":
        name = "Sam Sweet"
    if name == "Nivedha Vasanth":
        name = "Nivedha Prasanth"
    if name == "Gabriel Visotsky":
        name = "Gabe Visotsky"
    if name == "Maddison Wang":
        name = "Maddie Wang"
    if name == "Natalia Waszynska":
        name = "Natalia Wasynska"
    if name == "Mark Younan":
        name = "Mark Younnan"
    if name == "Youran Zhu":
        name = "Youran(Bill) Zhu"
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

math = xlrd.open_workbook("/Users/krishna/Documents/pt_attendance/FNL-Attendance_PeerTutor_SCIENCE.xlsx")
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

math = xlrd.open_workbook("/Users/krishna/Documents/pt_attendance/FNL-AT_PT_SS.xlsx")
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


math = xlrd.open_workbook("/Users/krishna/Documents/pt_attendance/FNL_OLK_CA_PT_AT.xlsx")
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


math = xlrd.open_workbook("/Users/krishna/Documents/pt_attendance/FNL-AT_PT_WL.xlsx")
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

math = xlrd.open_workbook("/Users/krishna/Documents/pt_attendance/FNL-AT_PT_GS.xlsx")
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

math = xlrd.open_workbook("/Users/krishna/Documents/pt_attendance/FNL_PIO_CA_PT_AT.xlsx")
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

math = xlrd.open_workbook("/Users/krishna/Documents/pt_attendance/DW Attendance tracking for teachers.xlsx")
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

math = xlrd.open_workbook("/Users/krishna/Documents/pt_attendance/FNL-AT_PT_CS.xlsx")
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

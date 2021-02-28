import pandas as pd
import xlrd 
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
responses = xlrd.open_workbook("insert_excelsheet_name.xlsx").sheet_by_name("Form Responses 1")
tutors = []
extra_shift_tutors = []
monday_col = 11
extra_shift_col = 18
for x in range(1, responses.nrows):
    name = responses.cell(x, 3).value + " " + responses.cell(x, 2).value
    level = len(responses.cell(x, 10).value.split(","))
    if responses.cell(x, 6).value == "CommArts":
        tutors.append({"name": name, "level": level, "monday": str(responses.cell(x, monday_col).value).split(", "), "tuesday": str(responses.cell(x, monday_col+1).value).split(", "), "wednesday": str(responses.cell(x, monday_col+2).value).split(", "), "thursday": str(responses.cell(x, monday_col+3).value).split(", "), "friday": str(responses.cell(x, monday_col+4).value).split(", ")})
        if responses.cell(x, 18).value != 1 and responses.cell(x, extra_shift_col).value != 1.0 and responses.cell(x, extra_shift_col).value != '1.0' and responses.cell(x, extra_shift_col).value != '1' and responses.cell(x,extra_shift_col).value != '':
            extra_shift_tutors.append({"name": name, "level": level, "monday": str(responses.cell(x, monday_col).value).split(", "), "tuesday": str(responses.cell(x, monday_col+1).value).split(", "), "wednesday": str(responses.cell(x, monday_col+2).value).split(", "), "thursday": str(responses.cell(x, monday_col+3).value).split(", "), "friday": str(responses.cell(x, monday_col+4).value).split(", "), 'shifts': responses.cell(x, extra_shift_col).value})

shifts = {"7:30-8:30": [], "1st": [], "2nd": [], "3rd": [], "4th": [], "5th": [], "6th": [], "7th": [], "8th": [], "3:30-4:30": [], "4:30-5:30 (Only for math)": []}
schedule = {"monday": {"7:30-8:30": [], "1st": [], "2nd": [], "3rd": [], "4th": [], "5th": [], "6th": [], "7th": [], "8th": [], "3:30-4:30": [], "4:30-5:30 (Only for math)": []}, "tuesday": {"7:30-8:30": [], "1st": [], "2nd": [], "3rd": [], "4th": [], "5th": [], "6th": [], "7th": [], "8th": [], "3:30-4:30": [], "4:30-5:30 (Only for math)": []}, "wednesday": {"7:30-8:30": [], "1st": [], "2nd": [], "3rd": [], "4th": [], "5th": [], "6th": [], "7th": [], "8th": [], "3:30-4:30": [], "4:30-5:30 (Only for math)": []}, "thursday": {"7:30-8:30": [], "1st": [], "2nd": [], "3rd": [], "4th": [], "5th": [], "6th": [], "7th": [], "8th": [], "3:30-4:30": [], "4:30-5:30 (Only for math)": []}, "friday": {"7:30-8:30": [], "1st": [], "2nd": [], "3rd": [], "4th": [], "5th": [], "6th": [], "7th": [], "8th": [], "3:30-4:30": [], "4:30-5:30 (Only for math)": []}}
monday_pop_times = ["7:30-8:30", "3:30-4:30"]

def convertTimes (times):
    for x in range(len(times)):
        if times[x] == "1" or times[x] == "1.0":
            times[x] = '1st'
        if times[x] == "2" or times[x] == "2.0":
            times[x] = '2nd'
        if times[x] == "3" or times[x] == "3.0":
            times[x] = '3rd'
        if times[x] == "4" or times[x] == "4.0":
            times[x] = '4th'
        if times[x] == "5" or times[x] == "5.0":
            times[x] = '5th'
        if times[x] == '6' or times[x] == '6.0':
            times[x] = '6th'
        if times[x] == "7" or times[x] == '7.0':
            times[x] = '7th'
        if times[x] == '8' or times[x] == '8.0':
            times[x] = '8th'
    return times
def calculate_num_shifts(students):
    for t in students:
        t["num_shifts"] = 0
        if t['monday'][0] == '':
            pass
        else:
            t["num_shifts"] += len(t["monday"])
        if t['tuesday'][0] == '':
            pass
        else:
            t["num_shifts"] += len(t["tuesday"])
        if t['wednesday'][0] == '':
            pass
        else:
            t["num_shifts"] += len(t["wednesday"])
        if t['thursday'][0] == '':
            pass
        else:
            t["num_shifts"] += len(t["thursday"])
        if t['friday'][0] == '':
            pass
        else:
            t["num_shifts"] += len(t["friday"])
    return students
def sort_periods(sched, students, days):
    for d in days:
        tutor_num = 0
        while tutor_num < len(students):
            tutor = students[tutor_num]
            times = convertTimes(tutor[d])
            tutor_num+=1
            for period in list(shifts.keys())[1:9]:
                already_in_shift = list(i["name"] for i in sched[d][period]).count(tutor["name"])
                if times.count(period)>0 and already_in_shift == 0 and len(sched[d][period]) < 3:
                    sched[d][period].append(tutor)
                    students.remove(tutor)
                    tutor_num-=1
                    break 
    return sched, students
def sort_morning_afternoon(schedule, tutors, days):
    for d in days:
        total_level = 0
        for x in monday_pop_times:
            total_level = 0
            tutor_num = 0
            while tutor_num < len(tutors):
                tutor = tutors[tutor_num]
                times = convertTimes(tutor[d])
                tutor_num+= 1
                already_in_shift = list(i["name"] for i in schedule[d][x]).count(tutor["name"])
                #print("{0} appears {1} times in this shift".format(tutor["name"], already_in_shift))
                if times.count(x)>0 and already_in_shift == 0 and (len(schedule[d][x]) < 5 or ((d == 'friday' or d == 'thursday') and len(schedule[d][x]) < 5)):
                        schedule[d][x].append(tutor)
                        total_level += tutor['level']
                        #print('{0} to {1} at {2}'.format(tutor['name'], d, x))
                        tutors.remove(tutor)
                        tutor_num -= 1
    return schedule, tutors

tutors = sorted(calculate_num_shifts(tutors), key = lambda i: i['num_shifts'])
days = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday']
schedule, tutors = sort_periods(schedule, tutors, days)
schedule, tutors = sort_morning_afternoon(schedule, tutors, days)


extra_shift_tutors = sorted(calculate_num_shifts(extra_shift_tutors), key = lambda i: i['num_shifts'])
days = ['thursday', 'wednesday', 'friday', 'tuesday', 'monday']
schedule, extra_shift_tutors = sort_periods(schedule, extra_shift_tutors, days)
schedule, extra_shift_tutors = sort_morning_afternoon(schedule, extra_shift_tutors, days)
days = ['friday', 'wednesday', 'tuesday', 'monday', 'thursday']
schedule, extra_shift_tutors = sort_periods(schedule, extra_shift_tutors, days)
schedule, extra_shift_tutors = sort_morning_afternoon(schedule, extra_shift_tutors, days)
days = ['wednesday', 'tuesday', 'tuesday', 'monday', 'friday']
schedule, extra_shift_tutors = sort_periods(schedule, extra_shift_tutors, days)

for x in schedule.keys():
    print((str(x)+"\n"))
    for y in schedule[x].keys():
        names = [i["name"] for i in schedule[x][y]]
        print(str(y) + ": " + str(names))
    print("\n\n")
print(tutors)
print(extra_shift_tutors)
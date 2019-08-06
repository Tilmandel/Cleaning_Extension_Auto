import json
from csv import DictReader as DR
import xlsxwriter
import xlrd,csv

import time

users_detail_fusion_report = {}
user_detail_extension_file = {}
cleaned_extension = {}
def fusion_report():
    movefile= "all wro_csv.csv"
    move_file = open(movefile, "rb")
    move_read = DR(move_file)
    for i in move_read:
        users_detail_fusion_report[i["GPN"]]=i['Imie i Nazwisko'], i['Imie'],i['Nazwisko']
def extension_data_base():
    movefile= "extension_CSV.csv"
    move_file = open(movefile, "rb")
    move_read = DR(move_file)
    for i in move_read:
        if len(i["Numer"]) ==4 or len(i["Numer"]) ==7 and\
            i['Osoba'] != '' and i['GPN'] != ''\
            or i['Osoba'] == '' and i['GPN'] == ''\
            or i['Osoba'] != '' and i['GPN'] == ''\
            or i['Osoba'] == '' and i['GPN'] != '':
            if len("Numer") == 4:
                user_detail_extension_file["475"+i['Numer']] = i['Osoba'],i['GPN']
            else:
                user_detail_extension_file[i['Numer']] = i['Osoba'], i['GPN']
def csv_from_excel():
    wb = xlrd.open_workbook('S:\UCC\Phone Extension Wroclaw.xlsx')
    sh = wb.sheet_by_name('My sheet')
    your_csv_file = open('Phone_extension_CSV.csv', 'w')
    wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)

    for rownum in range(sh.nrows):
        wr.writerow(sh.row_values(rownum))

    your_csv_file.close()
def phone_extension_db():
    movefile= "Phone_extension_CSV.csv"
    move_file = open(movefile, "rb")
    move_read = DR(move_file)
    for i in move_read:
        user_detail_extension_file[i["Extension"].replace(".0","")] = i['GPN'].replace(".0",""),i["Name & LastName"]
def _write_cleaned_to_xlsx():
    workbook = xlsxwriter.Workbook('S:\UCC\Phone Extension Wroclaw.xlsx')
    worksheet = workbook.add_worksheet("My sheet")
    bold = workbook.add_format({'bold': True})
    worksheet.set_column(0, 0, 6)
    worksheet.set_column(0, 1, 6)
    worksheet.set_column(0, 2, 25)
    worksheet.write(0, 0, "Extension", bold)
    worksheet.write(0, 1, "GPN", bold)
    worksheet.write(0, 2, "Name & LastName", bold)

    row = 1
    for x in cleaned_extension:
        if len(x) ==4:
            numer = int("475"+x)
        else:
            numer = int(x)
        if "Free" in cleaned_extension[x][1]:
            name = cleaned_extension[x][1]
            worksheet.write(row, 2, name)
        if "Free" not in cleaned_extension[x][1]:
            name = cleaned_extension[x][1]
            worksheet.write(row, 2, name)

        try:
            gpn = int(cleaned_extension[x][0])
        except ValueError:
            gpn = cleaned_extension[x][0]
        worksheet.write(row, 1, gpn)
        worksheet.write(row, 0, numer)
        row += 1
    workbook.close()
def _gpn_cross_name_Hotlines_etc():
    to_be_removed = []
    for x in users_detail_fusion_report.keys():
        for y in user_detail_extension_file:
            if x == user_detail_extension_file[y][1]:
                name = users_detail_fusion_report[x][1] + " " + users_detail_fusion_report[x][2]
                numer = y
                gpn = user_detail_extension_file[y][1]
                cleaned_extension[numer] = gpn, name
                to_be_removed.append(y)
    for i in to_be_removed:
        user_detail_extension_file.pop(i)

    to_be_removed = []
    for x in users_detail_fusion_report:
        for y in user_detail_extension_file:
            if users_detail_fusion_report[x][1] in user_detail_extension_file[y][0] and users_detail_fusion_report[x][2] in user_detail_extension_file[y][0]:
                name = users_detail_fusion_report[x][1] + " " + users_detail_fusion_report[x][2]
                numer = y
                gpn = x
                cleaned_extension[numer] = gpn, name
                to_be_removed.append(y)
    for i in to_be_removed:
        user_detail_extension_file.pop(i)

    to_be_removed = []
    for x in user_detail_extension_file:
        if "room" in user_detail_extension_file[x][0]:
            gpn = "Empty"

            cleaned_extension[x] = gpn, user_detail_extension_file[x][0]
            to_be_removed.append(x)
    for i in to_be_removed:
        user_detail_extension_file.pop(i)

    to_be_removed = []
    for x in user_detail_extension_file:
        if "line" in user_detail_extension_file[x][0] \
                or "Recep" in x \
                or "LINE" in x \
                or "CU" in x \
                or "CLM" in x \
                or "CISO" in x \
                or "Mgmt01" in x \
                or "Security" in x \
                or "Processing" in x \
                or "Wroclaw IT" in x:

            gpn = "Empty"
            cleaned_extension[x] = gpn, user_detail_extension_file[x][0]
            to_be_removed.append(x)
    for i in to_be_removed:
        user_detail_extension_file.pop(i)
    count = 0
    for x in user_detail_extension_file:
        name = "Free {}".format(count)
        gpn = "Empty"

        cleaned_extension[x] = gpn, name
        count += 1
def _crosse_check_new():
    to_be_removed = []
    for x in users_detail_fusion_report:
        for y in user_detail_extension_file:
            if x == user_detail_extension_file[y][0]:
                cleaned_extension[y] = user_detail_extension_file[y][0], user_detail_extension_file[y][1]
                to_be_removed.append(y)
    for i in to_be_removed:
        user_detail_extension_file.pop(i)
    to_be_removed = []
    numbers = []
    for i in user_detail_extension_file:
        try:
            if user_detail_extension_file[i][1].split()[0] == "Free":
                cleaned_extension[i] = user_detail_extension_file[i][0], user_detail_extension_file[i][1]
                numbers.append(int(user_detail_extension_file[i][1].split()[1]))
                to_be_removed.append(i)
        except IndexError:
            count =max(sorted(numbers))
            cleaned_extension[i] = "Empty", "Free {}".format(count)
            numbers.append(int(count))
            to_be_removed.append(i)

        if      "room" in user_detail_extension_file[i][1] \
                or "Recep" in user_detail_extension_file[i][1] \
                or "hot" in user_detail_extension_file[i][1] \
                or "Hot" in user_detail_extension_file[i][1] \
                or "CU" in user_detail_extension_file[i][1] \
                or "CLM" in user_detail_extension_file[i][1] \
                or "CISO" in user_detail_extension_file[i][1] \
                or "Mgmt01" in user_detail_extension_file[i][1] \
                or "Security" in user_detail_extension_file[i][1] \
                or "Processing" in user_detail_extension_file[i][1] \
                or "Wroclaw IT" in user_detail_extension_file[i][1] \
                or "HOT" in user_detail_extension_file[i][1]:
            cleaned_extension[i] = user_detail_extension_file[i][0], user_detail_extension_file[i][1]
            to_be_removed.append(i)
    for i in to_be_removed:
        user_detail_extension_file.pop(i)

    try:
        count = max(sorted(numbers))
    except TypeError:
        pass
    for i in user_detail_extension_file:
        count += 1
        name = "Free {}".format(count)
        gpn = "Empty"
        numer = i
        cleaned_extension[numer] = gpn, name
def _Free_numbers_count():
    count = 0
    for i in cleaned_extension:
        try:
            int(cleaned_extension[i][1].split()[1])
            count +=1
        except ValueError:
            continue
        except TypeError:
            continue
    return count


csv_from_excel()
fusion_report()
#=====Clean from first Files=====
#extension_data_base()
#_gpn_cross_name_Hotlines_etc()
#================================
phone_extension_db()
_crosse_check_new()
_write_cleaned_to_xlsx()

print ("="*30).center(20)
print "Summary".center(20)
print "Freded Numbers: {}".format(len(user_detail_extension_file)).center(20)
print "Hired Employee: {}".format(len(users_detail_fusion_report)).center(20)
print "All Number Count: {}".format(len(cleaned_extension)).center(20)
print "Total count of  Free number: {}".format(_Free_numbers_count()).center(20)
print ("="*30).center(20)
print user_detail_extension_file

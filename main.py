#!/usr/bin/python
__author__ = 'samirw'

import xlrd
import datetime
import smtplib
import time

# Email to send from
fromaddr = 'samirw@mit.edu'
username = 'swadhwania96@gmail.com'
password = 'rfwjgtzjrzuvewld'
birthday_chair = 'Amelia Bryan (aybryan@mit.edu)'

# Dates
now = datetime.datetime.now()
one_week = now + datetime.timedelta(days=7)
one_day = now + datetime.timedelta(days=1)

# Columns
name_col = 1
kerberos_col = 2
birthday_col = 6
candy_col = 7

# Filename
data_file = 'download.xlsx'
b5_bday_sheet_name = 'B5 Birthdays'
b5_roster_sheet_name = 'B5 Roster'

xl_workbook = xlrd.open_workbook(data_file)

def main():
    b5_bday_sheet = xl_workbook.sheet_by_name(b5_bday_sheet_name)

    for row in range(b5_bday_sheet.nrows):
        date = convert_date(b5_bday_sheet.row(row))
        if date:
            find_suite(b5_bday_sheet.row(row), date)
        else:
            pass

def convert_date(row):
    try:
        xl_date = row[birthday_col].value
        bday_date = xlrd.xldate_as_tuple(xl_date, xl_workbook.datemode)

        if check_date(bday_date):
            return bday_date
    except:
        pass

    return False

def check_date(bday):
    if one_week.month == bday[1] or one_day.month == bday[1]:
        if one_week.day == bday[2] or one_day.day == bday[2]:
            return True
    return False

def find_suite(row, date):

    def find_others(suite):
        others = []
        for row in range(b5_roster_sheet.nrows):
            suite_check = str(b5_roster_sheet.row(row)[0])
            suite_check = suite_check[-5:-2]
            if suite_check == suite:
                others.append(str(b5_roster_sheet.row(row)[2]))

        return others

    b5_roster_sheet = xl_workbook.sheet_by_name(b5_roster_sheet_name)
    kerb = str(row[kerberos_col])
    stu_name = str(row[name_col])[7:-1]
    stu_candy = str(row[candy_col])[7:-1]
    bday_date = str(date[1]) + '/' + str(date[2])


    for row in range(b5_roster_sheet.nrows):
        name = str(b5_roster_sheet.row(row)[2])
        if name == kerb:
            suite = str(b5_roster_sheet.row(row)[0])
            suite = suite[-5:-2]
            suite_mates = find_others(suite)
            suite_mates.remove(kerb)

    emails = [s[7:-1] + "@mit.edu" for s in suite_mates]
    emails.append('burton5-exec@mit.edu')

    send_email(stu_name, bday_date, stu_candy, emails)

def send_email(name, birthday, candy, emails):
    message = """Subject: Upcoming Birthday

Hello!

This is an automated message to remind you that on %(birthday)s, %(name)s will be celebrating their birthday! Their favorite candy is: %(candy)s. Remember, it is the responsibility of the suite to make/get something to celebrate. If you do not have ingredients or would like to buy something, please email %(chair)s. Have a great day!

Best wishes,
B5 Exec
""" % {"birthday" : birthday, "name" : name, "candy" : candy, "chair" : birthday_chair}

    server = smtplib.SMTP('smtp.gmail.com:587')
    server.starttls()
    server.login(username,password)
    server.sendmail(fromaddr, emails, message)
    server.quit()

def test_email():
    message = "Test 1"

    server = smtplib.SMTP('smtp.gmail.com:587')
    server.starttls()
    server.login(username,password)
    server.sendmail(fromaddr, "swadhwania96@gmail.com", message)
    server.quit()

if __name__ == "__main__":
    main()

#!/usr/bin/python
__author__ = 'samirw'

import xlrd
import datetime
import smtplib
import time

# Email to send from
fromaddr = 'babcockj@mit.edu'
username = 'jobabcock95@gmail.com'
password = 'pxolygbedozsiebx'
birthday_chair = 'Joe Babcock (babcockj@mit.edu)'

# Dates
now = datetime.datetime.now()
one_week = now + datetime.timedelta(days=7)
one_day = now + datetime.timedelta(days=1)

summer_birthday = False
# Columns
suite_col = 0
name_col = 1
kerberos_col = 2
birthday_col = 4


# Filename
data_file = 'Burton5_Roster_and_Birthdays.xlsx'
b5_bday_sheet_name = 'Fall2016'

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

    if bday[1] in [6,7,8]:
        check_mon = bday[1] - 6
        summer_birthday = True
        summer_bday=(bday[0],check_mon,bday[2])
    else:
        check_mon = bday[1]
    check_day = bday[2]

    if (one_week.month,one_week.day) == (check_mon, check_day) or (one_day.month,one_day.day) == (check_mon, check_day):
            return True
    return False

def find_suite(row, date):
    emails = []

    b5_sheet = xl_workbook.sheet_by_name(b5_bday_sheet_name)
    room = str(row[suite_col].value)
    suite = room[:-1]
    for i in range(b5_sheet.nrows):
        row_temp = b5_sheet.row(i)
        room_temp = str(row_temp[suite_col].value)
        suite_temp = room_temp[:-1]
        if suite_temp == suite and room_temp != room:
            emails.append(str(row_temp[kerberos_col].value))

    emails = [x + "@mit.edu" for x in emails]
    emails.append('burton5-exec@mit.edu')

    name = str(row[name_col].value)
    date = str(date[1]) + '/' + str(date[2])

    if summer_birthday == True:
        send_email_summer(name, summer_bday, emails)
        summer_birthday=False
    else:
        send_email(name, date, emails)

def send_email(name, birthday, emails):
    message = """Subject: Upcoming Birthday

Hello!

This is an automated message to remind you that on %(birthday)s, %(name)s will be celebrating their birthday! Remember, it is the responsibility of the suite to make/get something to celebrate. If you do not have ingredients or would like to buy something, please email %(chair)s. Have a great day!

Best wishes,
B5 Exec
""" % {"birthday" : birthday, "name" : name, "chair" : birthday_chair}

    server = smtplib.SMTP('smtp.gmail.com:587')
    server.starttls()
    server.login(username,password)
    server.sendmail(fromaddr, emails, message)
    server.quit()

def send_email_summer(name, birthday, emails):
    message = """Subject: Upcoming Birthday

Hello!

This is an automated message to remind you that on %(birthday)s, %(name)s will be celebrating their half sbirthday! Remember, it is the responsibility of the suite to make/get something to celebrate. If you do not have ingredients or would like to buy something, please email %(chair)s. Have a great day!

Best wishes,
B5 Exec
""" % {"birthday" : birthday, "name" : name, "chair" : birthday_chair}

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
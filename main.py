#!/usr/bin/python3.6
import sys
import xlrd
import xlwt
from xlwt import Workbook
from email_validator import validate_email, EmailNotValidError


def email_splitter(email):
    username = email.split('@')[0]
    first_name = username.split('.')[0]
    last_name = username.split('.')[1]
    return first_name, last_name


loc = sys.argv[1]
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
wb1 = Workbook()
sheet1 = wb1.add_sheet('Sheet 1')
style = xlwt.easyxf('font: bold 1, color blue;')
sheet1.write(0, 0, 'First Name', style)
sheet1.write(0, 1, 'Last Name', style)
try:
    for i in range(1, sheet.nrows):
        email = sheet.cell_value(i, 2)
        valid = validate_email(email)
        email = valid.email
        first_name, last_name = email_splitter(email)
        sheet1.write(i, 0, first_name)
        sheet1.write(i, 1, last_name)

    wb1.save('Sheet_1.xls')

except EmailNotValidError as e:
    print(str(e))

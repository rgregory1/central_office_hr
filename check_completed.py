import pygsheets
from pprint import pprint
# import os
import yagmail
import credentials
import datetime

# setup gmail link
gmail_user = credentials.gmail_user
gmail_password = credentials.gmail_password
yag = yagmail.SMTP(gmail_user, gmail_password)


# get timestamp for log
temp_timestamp = str(datetime.datetime.now())
print(temp_timestamp)
print('\n')


# open new staff process sheet
gc = pygsheets.authorize(outh_file='client_secret.json')
workbook = gc.open_by_key('1KWLOYV7wQjEaD0A107gZlivZ3sr8OeOcDP3OjVOSX6E')
MasterList = workbook.worksheet_by_title("MasterList")

# grab info from the master list
master_list_matrix_raw = MasterList.get_all_values(returnas='matrix')
# remove blank rows
master_list_matrix = [x for x in master_list_matrix_raw if x[0] != '']


# create list of dictionary keys
dict_key_list = [x for x in master_list_matrix[0] if x != '']
# remove headers from matrix
master_list_matrix = master_list_matrix[1:]

# initialize master list
master_list_data = []

# put cell_matrix list of lists into a list of dictionaries
for count, row in enumerate(master_list_matrix):
    line_dict = dict(zip(dict_key_list, row))
    # add count so I can add x to appropriate row later
    line_dict['row_number'] = count + 1
    master_list_data.append(line_dict)

pprint(master_list_data)

# initialize final strings
final_bonnie_todo = ''
final_jeri_todo = ''
final_pierrette_todo = ''
final_michelle_todo = ''


# begin loop loooing for incomplete staff members
for staff in master_list_data:
    if staff['Status'] == 'Not Complete':
        # print(master_list_data[staff]['Staff Name'])
        this_staff_sheet = workbook.worksheet_by_title(staff['Staff Name'])
        # this_staff_matrix = this_staff_sheet.get_all_values(returnas='matrix')

        # begin bonnie email notifications
        bonnie_range = this_staff_sheet.range('A12:B16')
        bonnie_todo = ''

        for count, line in enumerate(bonnie_range):
            # print(this_staff_data[number])
            if bonnie_range[count][0].value == '':
                bonnie_todo = bonnie_todo + bonnie_range[count][1].value + '\n'
        if bonnie_todo != '':
            final_bonnie_todo = final_bonnie_todo + \
                staff['Staff Name'] + '\n \n' + bonnie_todo + '\n\n'

        # begin jeri email notifications
        jeri_range = this_staff_sheet.range('A21:B25')
        jeri_todo = ''

        for count, line in enumerate(jeri_range):
            # print(this_staff_data[number])
            if jeri_range[count][0].value == '':
                jeri_todo = jeri_todo + jeri_range[count][1].value + '\n'
        if jeri_todo != '':
            final_jeri_todo = final_jeri_todo + \
                staff['Staff Name'] + '\n \n' + jeri_todo + '\n\n'

        # begin pierrette email notifications
        pierrette_range = this_staff_sheet.range('A30:B31')
        pierrette_todo = ''

        for count, line in enumerate(pierrette_range):
            # print(this_staff_data[number])
            if pierrette_range[count][0].value == '':
                pierrette_todo = pierrette_todo + \
                    pierrette_range[count][1].value + '\n'
        if pierrette_todo != '':
            final_pierrette_todo = final_pierrette_todo + \
                staff['Staff Name'] + '\n \n' + pierrette_todo + '\n\n'

        # begin michelle email notifications
        michelle_range = this_staff_sheet.range('A36:B42')
        michelle_todo = ''

        for count, line in enumerate(michelle_range):
            # print(this_staff_data[number])
            if michelle_range[count][0].value == '':
                michelle_todo = michelle_todo + \
                    michelle_range[count][1].value + '\n'
        if michelle_todo != '':
            final_michelle_todo = final_michelle_todo + \
                staff['Staff Name'] + '\n \n' + michelle_todo + '\n\n'


contents = 'This is your friendly weekly reminder of things to do for new staff members. \n \n \n'
contents2 = 'Due to your efficiency, there is actually nothing for you to do for new hires!'
html = '<a href="https://docs.google.com/spreadsheets/d/1KWLOYV7wQjEaD0A107gZlivZ3sr8OeOcDP3OjVOSX6E/edit#gid=0">New Staff Process spreadsheet</a>'


# bonnie emails
if final_bonnie_todo != '':
    yag.send(['bonnie.moulton@mvsdschools.org', 'russell.gregory@mvsdschools.org'],
             'New Staff Weekly Reminder', [contents, final_bonnie_todo, html])
else:
    yag.send('bonnie.moulton@mvsdschools.org',
             'New Staff Weekly Reminder', [contents, contents2, html])
print('bonnie email sent')

# jeri emails
if final_jeri_todo != '':
    yag.send('Jeri.Patterson@mvsdschools.org',
             'New Staff Weekly Reminder', [contents, final_jeri_todo, html])
else:
    yag.send('Jeri.Patterson@mvsdschools.org',
             'New Staff Weekly Reminder', [contents, contents2, html])
print('jeri email sent')

# pierrette emails
if final_pierrette_todo != '':
    yag.send('Pierrette.Bouchard@mvsdschools.org', 'New Staff Weekly Reminder', [
             contents, final_pierrette_todo, html])
else:
    yag.send('Pierrette.Bouchard@mvsdschools.org',
             'New Staff Weekly Reminder', [contents, contents2, html])
print('pierrette email sent')

# michelle emails
if final_michelle_todo != '':
    yag.send('Michelle.Stanley@mvsdschools.org', 'New Staff Weekly Reminder', [
             contents, final_michelle_todo, html])
else:
    yag.send('Michelle.Stanley@mvsdschools.org',
             'New Staff Weekly Reminder', [contents, contents2, html])
print('michelle email sent')

print('\n\nfinished')

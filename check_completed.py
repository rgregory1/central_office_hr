import pygsheets
from pprint import pprint
# import os
import yagmail
import credentials
import datetime

# setup gmail link
gmail_user = credentials.gmail_user
gmail_password = credentials.gmail_password
yag = yagmail.SMTP( gmail_user, gmail_password)


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
bonnie_todo = ''
jeri_todo = ''
pierrette_todo = ''
michelle_todo = ''



# # begin loop loooing for incomplete staff members
# for staff in master_list_data:
#     if master_list_data[staff]['Status'] == 'Not Complete':
#         # print(master_list_data[staff]['Staff Name'])
#         this_staff_sheet = workbook.worksheet_by_title(master_list_data[staff]['Staff Name'])
#         this_staff_matrix = this_staff_sheet.get_all_values(returnas='matrix')

#         counter = 1
#         this_staff_data = {}
#         new_line = {}

#         for line in this_staff_matrix:
#             this_line_data = {}
#             # this_line_data['row'] = counter
#             this_line_data['a'] = line[0]
#             this_line_data['b'] = line[1]
#             this_staff_data[counter] = this_line_data
#             counter = counter + 1
#         # pprint(this_staff_data)

#         # begin admin email notifications
#         admin_list = [12,13,14,15,16,17,18,19,20]
#         admin_todo = ''
#         for number in admin_list:
#             # print(this_staff_data[number])
#             if this_staff_data[number]['a'] == '':
#                 admin_todo = admin_todo + this_staff_data[number]['b'] + '\n'
#         if admin_todo != '':
#             final_admin_todo = final_admin_todo + master_list_data[staff]['Staff Name'] + '\n \n' + admin_todo + '\n\n'
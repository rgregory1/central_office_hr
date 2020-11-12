import pygsheets
import datetime
import credentials
import yagmail

print('\n' * 10)

is_new_term = False

# setup credentials for sending email
gmail_user = credentials.gmail_user
gmail_password = credentials.gmail_password
yag = yagmail.SMTP(gmail_user, gmail_password)

# get timestamp for log
temp_timestamp = str(datetime.datetime.now())
print(temp_timestamp)
print('\n')
print('checking new termination form entries')
print('\n')


def check_for_new_staff():
    print('Starting check of new termination form...')

    # open up google sheet to see if new termination have been added
    gc = pygsheets.authorize(outh_file='client_secret.json')
    initial_form_wb = gc.open_by_key(
        '1b3ktqopatkcrdLBbgAl3yTfRVKdvr_0yWAA_yNcNkaI')
    initial_form_sheet = initial_form_wb.worksheet_by_title("NewTermination")

    # download all data from sheet as cell_matrix
    cell_matrix = initial_form_sheet.get_all_values(returnas='matrix')
    # print(cell_matrix)

    # gather 'keys' for new dict from 1st row in sheet
    dict_key_list = [x for x in cell_matrix[0]]

    # initialize dict for data
    worksheet_data = []

    # put cell_matrix list of lists into a list of dictionaries
    for count, row in enumerate(cell_matrix):
        if row[9] == '':
            line_dict = dict(zip(dict_key_list, row))
            # add count so I can add x to appropriate row later
            line_dict['row_number'] = count + 1
            worksheet_data.append(line_dict)

    print('Check complete')

    return worksheet_data, initial_form_sheet


# check for new termination
worksheet_data, initial_form_sheet = check_for_new_staff()

# if new termination
# empty list
if len(worksheet_data) == 0:
    # empty list
    print("No new terminations found")

else:
    # list contains items
    print("New termination found")
    is_new_term = True

    # print(worksheet_data)

    # get blank copy of staff record sheet
    gc = pygsheets.authorize(outh_file='client_secret.json')
    fresh_copy_wkb = gc.open_by_key(
        '1pI4O0XZWHU2Jd7zL30dAtYhmCIe2GpqNf-DYT-326BU')
    fresh_copy_sheet = fresh_copy_wkb.worksheet_by_title("Original")

    # open CO New Termination Process google sheet
    staff_process_wkb = gc.open_by_key(
        '1CT2Xv3sOfQbi7HvLVwZWDKzY380QowGoOu4Cnx44MFo')
    master_list = staff_process_wkb.worksheet_by_title('MasterList')

    for staff in worksheet_data:

        staff_name = staff['First Name'] + ' ' + staff['Last Name']
        print(f'\nbegin creating sheet for {staff_name}')

        # create staff record sheet
        individual_record_sheet = staff_process_wkb.add_worksheet(
            staff_name, src_worksheet=fresh_copy_sheet)

        # move new sheet to first position
        individual_record_sheet.index = 1

        # add new staff members basic info to record sheet
        basic_info = individual_record_sheet.range('C2:C7')
        basic_info[0][0].value = staff_name
        basic_info[1][0].value = staff['School Location']
        basic_info[2][0].value = staff['Category']
        basic_info[3][0].value = staff['Position']
        basic_info[4][0].value = staff['Last Day of Employment']
        basic_info[5][0].value = staff['Retirement']

        # add new staff member to MasterList sheet

        # download all data from sheet as cell_matrix
        master_list_matrix = master_list.get_all_values(returnas='matrix')

        # find first empty row
        for count, row in enumerate(master_list_matrix):
            if row[0] == '':
                # print(count)
                new_staff_row_number = count + 1
                break
        start_range = 'A' + str(new_staff_row_number)
        end_range = 'D' + str(new_staff_row_number)
        new_staff_line = master_list.range(start_range + ':' + end_range)

        # populate first emtpy row with formulas from staff record sheet
        new_staff_line[0][0].value = staff_name
        new_staff_line[0][1].formula = "'" + staff_name + "'!C8"
        new_staff_line[0][2].formula = "'" + staff_name + "'!D12"
        new_staff_line[0][3].formula = "'" + staff_name + "'!D19"

        # mark new staff member as processed with X in column I
        mark_as_finished_cell = 'J' + str(staff['row_number'])
        initial_form_sheet.update_value(mark_as_finished_cell, 'X')

        print('finished with spreadsheet setup')

        # begin email notifications
        contents = 'A new staff termination, <b>' + staff_name + \
            '</b>, was added to the CO New Staff Termination spreadsheet, go and check it out. \n\n'
        html = '<a href="https://docs.google.com/spreadsheets/d/1CT2Xv3sOfQbi7HvLVwZWDKzY380QowGoOu4Cnx44MFo/edit#gid=2040843378">New Staff Termination spreadsheet</a>'
        yag.send(['russell.gregory@mvsdschools.org',
                  'bonnie.moulton@mvsdschools.org',
                  'Michelle.Stanley@mvsdschools.org',
                  ],
                 'New Termination',
                 [contents, html])
        print(f'sent notificaion email for {staff_name}\n\n')


if is_new_term == False:
    print('program comlpete, no new terminations')
else:
    end_timestamp = str(datetime.datetime.now())
    print(f'program complete at {end_timestamp}')

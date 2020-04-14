import pygsheets
import datetime

print('\n' * 10)



# get timestamp for log
temp_timestamp = str(datetime.datetime.now())
print(temp_timestamp)
print('checking new staff form entries')
print('\n')

def check_for_new_staff():
    # open up google sheet to see if new staff have been added
    gc = pygsheets.authorize(outh_file='client_secret.json')
    initial_form_wb = gc.open_by_key('1lRbvNLr4EJQ8pQco7MRooxF4ompKEAdCFehlaZRCYOY')
    initial_form_sheet = initial_form_wb.worksheet_by_title("NewStaff")


    # download all data from sheet as cell_matrix
    cell_matrix = initial_form_sheet.get_all_values(returnas='matrix')
    # print(cell_matrix)

    # gather 'keys' for new dict from 1st row in sheet
    dict_key_list = [x for x in cell_matrix[0]]

    # initialize dict for data
    worksheet_data = []
    
    # put cell_matrix list of lists into a list of dictionaries
    for count, row in enumerate(cell_matrix):
        if row[8] == '':
            line_dict = dict(zip(dict_key_list, row))
            # add count so I can add x to appropriate row later
            line_dict['row_number'] = count
            worksheet_data.append(line_dict)
         
    
    return worksheet_data, initial_form_sheet





# check for new staff
worksheet_data, initial_form_sheet = check_for_new_staff()

# if new staff
if len(worksheet_data) == 0: 
    print ("The list is Empty") 
else: 
    print ("The list is not empty") 

print(worksheet_data)

    # for row in cell_matrix:
    #     if row[8] == '':
    #         print('yehaw')


# for row in wks:
#     print(row)

# this_cell = wks.cell('A3')
# print(this_cell.value)

# cell_list = wks.range('A1:I2')
# # for row in cell_list:
# #     print(row[0].value)
# print(cell_list)

# admin_status = wks.cell('B10')
# admin_status.value = 'Done!'
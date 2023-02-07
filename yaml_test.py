import openpyxl, os
from time import time

# def find_correct_last_row(in_worksheet):
#     """
#     get sheet
#     return last row number
#     """
#     last_filled_row = 0
#     for index, row in enumerate(in_worksheet.values, start=1):
#         corrent_row_empty = row.count(None) == len(row)
#         if corrent_row_empty:
#             empty_row_successively += 1
#         else:
#             empty_row_successively = 0
#             last_filled_row = index
#         if empty_row_successively >= 5:
#             break
#     return last_filled_row

# file = 'school3.xlsx'
# workbook =  openpyxl.load_workbook(file, keep_vba=False, read_only=True)
# demension = workbook['5. Сведения о кадрах'].calculate_dimension()
# print(demension)
# workbook['5. Сведения о кадрах'].reset_dimensions()
# demension = workbook['5. Сведения о кадрах'].calculate_dimension()
# print(demension)
# worksheet = workbook['5. Сведения о кадрах']
# #print(worksheet.max_row)
# for iteration_worksheet in workbook.worksheets:
#     print("Current: ", iteration_worksheet.max_row, "Fact: ", find_correct_last_row(iteration_worksheet))
# workbook.close()

# write_workbook = openpyxl.load_workbook(file, read_only=True)
# write_worksheet = write_workbook['2. Сведения об обучающихся']
#new_workbook = openpyxl.Workbook(write_only=True)
#new_workbook.create_sheet(title="copy_sheet")
#new_workbook.save(time(),'.xlsx')
#help(write_worksheet)
#help(write_worksheet.values)
#help(write_worksheet.cell(2,2).style_array)
#help(write_worksheet.iter_rows)
# for item in write_worksheet.iter_rows(3):
#     print(item)
#     #help(item)
#     break
# write_workbook.close() 

# def remove_exscess_row(workbook):
#     for iteration_worksheet in workbook.worksheets:
#         first_empty_row = find_correct_last_row(iteration_worksheet) + 1
#         max_row = iteration_worksheet.max_row
#         iteration_worksheet.delete_rows(first_empty_row, max_row-first_empty_row+1)


# wb = openpyxl.load_workbook('school3.xlsx')
# remove_exscess_row(wb)
# wb.save('school3.xlsx')
# wb.close()

file_list = ['C:\\Users\\alist\\OneDrive\\Документы\\osokin\\monitoring\\example.xlsx', 
             'C:\\Users\\alist\\OneDrive\\Документы\\osokin\\monitoring\\school3.xlsx', 
             'C:\\Users\\alist\\OneDrive\\Документы\\osokin\\monitoring\\babaevo\\school1.xlsx', 
             'C:\\Users\\alist\\OneDrive\\Документы\\osokin\\monitoring\\babaevo\\school2.xlsx', 
             'C:\\Users\\alist\\OneDrive\\Документы\\osokin\\monitoring\\babaevo\\school3.xlsx', 
             'C:\\Users\\alist\\OneDrive\\Документы\\osokin\\monitoring\\chagoda\\school2.xlsx']

new_list  = [os.path.split(file) for file in file_list]
print(new_list)
print(os.path.abspath(file_list[0]))
print(os.path.basename(file_list[0]))
#print(os.path.commonpath(file_list[0]))
print(os.path.dirname(file_list[0]))
path_list = sorted(list({ os.path.dirname(file) for file in file_list }), key=len, reverse=True)
print(path_list)
for path in path_list:
    
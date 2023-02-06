import openpyxl

def find_correct_last_row(in_worksheet):
    """
    get sheet
    return last row number
    """
    last_filled_row = 0
    for index, row in enumerate(in_worksheet.values):
        corrent_row_empty = row.count(None) == len(row)
        if corrent_row_empty:
            empty_row_successively += 1
        else:
            empty_row_successively = 0
            last_filled_row = index
        if empty_row_successively >= 5:
            break
    return last_filled_row

file = 'C:\\Users\\alist\\OneDrive\\Документы\\osokin\\monitoring\\babaevo\\school3.xlsx'
workbook =  openpyxl.load_workbook(file, keep_vba=False, read_only=True)
demension = workbook['5. Сведения о кадрах'].calculate_dimension()
print(demension)
# workbook['5. Сведения о кадрах'].reset_dimensions()
# demension = workbook['5. Сведения о кадрах'].calculate_dimension()
# print(demension)
worksheet = workbook['5. Сведения о кадрах']
#print(worksheet.max_row)
for iteration_worksheet in workbook.worksheets:
    print("Current: ", iteration_worksheet.max_row, "Fact: ", find_correct_last_row(iteration_worksheet))
workbook.close()

write_workbook = openpyxl.load_workbook(file, read_only=True)
write_worksheet = write_workbook['5. Сведения о кадрах']
help(write_worksheet.values)
write_workbook.close()
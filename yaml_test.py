import openpyxl

file = 'C:\\Users\\alist\\OneDrive\\Документы\\osokin\\monitoring\\babaevo\\school3.xlsx'
workbook =  openpyxl.load_workbook(file, keep_vba=False, read_only=True)
# demension = workbook['5. Сведения о кадрах'].calculate_dimension()
# print(demension)
# workbook['5. Сведения о кадрах'].reset_dimensions()
# demension = workbook['5. Сведения о кадрах'].calculate_dimension()
# print(demension)
worksheet = workbook['5. Сведения о кадрах']
for i in worksheet.values:
    print(i)
    if i.count(None) == len(i):
        break
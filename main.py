#
# Принцип работы программы:
#   1. смотрим какие есть папки с файлами, создаём списки и передаём дальше
#   2. проверяем файлы в папке на то устраивают они нас или нет, выкидываем ненужные
#   3. создаём файл  консолидирующие каждую папку
#   4. заполняем файл данными, формулами, либо обычной суммой, 
#   5. .... далее пока дожить надо
#
import os, openpyxl, joblib, time, yaml

start = time.time()

logs = open("logs.txt", "a")
logs.truncate(0)


def getXlsxList(dir):
    #
    # input directory outpoot list all xlsx files
    #
    file_list = list()
    for address, dirs, files in os.walk(dir):
        for name in files:
            if name.endswith('.xlsx'):
                file_list.append(os.path.join(address,name))
    return file_list

def normalize(none):
    pass

def checkFiles(all_file, *args):
    #
    # input list of xlsx files and filter functions
    # outut current files list
    #
    new_list = all_file
    for filter_function in args:
        new_list = list(filter(filter_function, new_list))
    return new_list

def filterHaveSheet(file_path):
    #
    # input file path, then open file and check
    # all sheet have in workbook
    # output boolean
    #
    needSheet = ['Титульный лист',
                 '1. Сведения об ОО',
                 '2. Сведения об обучающихся',
                 '3. Сведения о режиме работы ГПД',
                 '4. Сведения о помещениях',
                 '5. Сведения о кадрах',
                 '6. Финансирование']
    wb = openpyxl.load_workbook(file_path, keep_vba=False, read_only=True)
    if set(needSheet).issubset(wb.sheetnames):
        wb.close()
        return True
    else:
        logs.write(f'\nВ файле отсутствуют нужные страницы: {file_path}')
        wb.close()
        return False



# 1.
file_list = getXlsxList(os.getcwd())
logs.write(f"\nНайдены следующие файлы: {file_list}")
# 2.
filteredExcelFileList = checkFiles(file_list, filterHaveSheet)
logs.write(f"\nБудут использоваться файлы {filteredExcelFileList}")
print(filteredExcelFileList)
# 3.


logs.close # close logs file
#abra = input("press any button")
end = time.time()

print(end - start)
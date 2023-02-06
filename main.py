#
# Принцип работы программы:
#   1. смотрим какие есть папки с файлами, создаём списки и передаём дальше
#   2. проверяем файлы в папке на то устраивают они нас или нет, выкидываем ненужные
#   3. создаём файл  консолидирующие каждую папку
#   4. заполняем файл данными, формулами, либо обычной суммой, 
#   5. .... далее пока дожить надо
#
import os, openpyxl, time, yaml


def read_config(configname):
    """
    Read configuration yaml file
    """
    with open(configname, "r", encoding="utf-8") as conf_file:
        return yaml.safe_load(conf_file)
    
def get_xlsx_list(dir):
    """
    input directory outpoot list all xlsx files
    """
    file_list = list()
    for address, dirs, files in os.walk(dir):
        for name in files:
            if name.endswith('.xlsx'):
                file_list.append(os.path.join(address,name))
    return file_list

def find_last_row(workseet):
    """
    show last row
    """
    pass

def find_correct_last_row(worksheet):
    """
    get sheet
    return last row number
    """
    empty_row_successively = 0
    last_filled_row = 0
    for index, row in enumerate(worksheet.values):
        corrent_row_empty = row.count(None) == len(row)
        if corrent_row_empty:
            empty_row_successively += 1
        else:
            empty_row_successively = 0
            last_filled_row = index
        if empty_row_successively >= 5:
            break
    return last_filled_row

def normalize():
    """
    normalize excel file
    """
    pass

def check_files(all_file, *args):
    """
    input list of xlsx files and filter functions
    outut current files list
    """
    new_list = all_file
    for filter_function in args:
        new_list = list(filter(filter_function, new_list))
    return new_list

def filter_have_sheet(file_path):
    """
    input file path, then open file and check
    all sheet have in workbook
    output boolean
    """
    need_sheet = configuration['filter']['sheet'] if configuration['filter']['sheet'] else []
    work_book = openpyxl.load_workbook(file_path, keep_vba=False, read_only=True)
    if set(need_sheet).issubset(work_book.sheetnames):
        work_book.close()
        return True
    #logs.write(f'\nВ файле отсутствуют нужные страницы: {file_path}')
    work_book.close()
    return False

configuration = read_config("configuration.yaml")

def main():
    """
    Empty documentation
    """
    # Debug
    start = time.time()
    # 0. Читаем настройки программы и открывает логи для занесения данных
    logs = open("logs.txt", "a", encoding="utf-8")
    logs.truncate(0)
    # 1. Смотрим все файлы в указанной дериктории
    file_list = get_xlsx_list(configuration['folder'] != None if configuration['folder'] else os.getcwd())
    logs.write(f"\nНайдены следующие файлы: {file_list}")
    # 2.
    filteredExcelFileList = check_files(file_list, filter_have_sheet)
    logs.write(f"\nБудут использоваться файлы {filteredExcelFileList}")
    print(filteredExcelFileList)
    # 3.


    logs.close # close logs file
    #abra = input("press any button")
    end = time.time()

    print(end - start)

if __name__ == "__main__":
    main()
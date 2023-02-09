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
    for index, row in enumerate(worksheet.values, start=1):
        corrent_row_empty = row.count(None) == len(row)
        if corrent_row_empty:
            empty_row_successively += 1
        else:
            empty_row_successively = 0
            last_filled_row = index
        if empty_row_successively >= 5:
            break
    return last_filled_row

def remove_exscess_row(workbook):
    """
    Input openpyxl.workbook
    then delete all excess row
    dont save document
    """
    for iteration_worksheet in workbook.worksheets:
        first_empty_row = find_correct_last_row(iteration_worksheet) + 1
        max_row = iteration_worksheet.max_row
        iteration_worksheet.delete_rows(first_empty_row, max_row-first_empty_row+1)

def cell_strip_space(workbook):
    """
    Cleare all empty cell, replace them None value
    """
    for iteration_worksheet in workbook:
        for row in iteration_worksheet.rows:
            for cell in row:
                if isinstance(cell.value, str):
                    if len(cell.value.strip()) == 0:
                        cell.value = None
                    

def normalize(file_list):
    """
    normalize excel file
    """
    for file in file_list:
        if os.stat(file).st_size / (1024*1024) > 3:
            wb = openpyxl.load_workbook(file)
            remove_exscess_row(wb)
            wb.save(file)
            wb.close()
        wb = openpyxl.load_workbook(file)
        cell_strip_space(wb)
        wb.save(file)
        wb.close()
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

def dialog_menu(value):
    """
    """
def get_directory_sorted_list(dir):
    """
    Take list of currect file
    return sorted list dirctory
    """
    path_list = sorted(list({ os.path.dirname(file) for file in dir }), key=len, reverse=True)
    return path_list
def get_current_file_from_directory(file_list, dir_name):
    """
    return all needed file to work
    """
    work_file_list = list()
    for file in file_list:
        if os.path.split(file)[0] == dir_name:
            work_file_list.append(file)
    return work_file_list

def create_needed_sheet(workbook_name, sheet_list):
    """
    Create page if then not have
    """
    book = openpyxl.load_workbook(workbook_name)
    for sheet_name in sheet_list:
        if not hasattr(book, sheet_name):
            book.create_sheet(sheet_name)
    del(book['Sheet'])
    book.save(workbook_name)
    book.close()

def create_needed_sheets(workbook_list, sheet_list):
    """
    """
    pass

def file_generator(file_list):
    """
    Create file by list
    """
    created_file = list()
    for file in file_list:
        if not os.path.exists(f'{file}.xlsx'):
            new_book = openpyxl.Workbook()
            new_book.save(f'{file}.xlsx')
            new_book.close()
            created_file.append(f'{file}.xlsx')
    return created_file
def create_sheet_content(destination_file, source_files):
    """
    """
    dest_file = openpyxl.load_workbook(destination_file)
    exam_file = openpyxl.load_workbook('example.xlsx')
    for work_sheet in dest_file.worksheets:
        for row in exam_file[work_sheet.title].rows:
            work_sheet.append([add_data_title_sheet(cell, source_files) if work_sheet.title == "Титульный лист" else add_data_to_cell(cell, source_files) for cell in row])
    dest_file.save(destination_file)
    dest_file.close()
    exam_file.close()

def make_sheet_style(destination_file):
    """
    """
    dest_file = openpyxl.load_workbook(destination_file)
    exam_file = openpyxl.load_workbook('example.xlsx')
    for work_sheet in dest_file.worksheets:
        for row in exam_file[work_sheet.title].rows:
            for cell in row:
                if cell.has_style:
                    pass
    dest_file.save(destination_file)
    dest_file.close()
    exam_file.close()
def add_data_title_sheet(cell, source_files):
    """
    """
    data_cell = None
    if isinstance(cell.value, str):
        data_cell = cell.value
    else:
        data_cell = "="
        page_title = cell.parent.title
        for address in source_files:
            if data_cell[-1] != "=":
                data_cell += "&\" \"&"
            data_cell += "\'"
            data_cell += os.path.dirname(address)
            data_cell += os.sep
            data_cell += "["
            data_cell += os.path.basename(address)
            data_cell += "]"
            data_cell += page_title
            data_cell += "\'"
            data_cell += "!"
            data_cell += cell.coordinate
    return data_cell



def add_data_to_cell(cell, source_files):
    """
    """
    data_cell = None
    if isinstance(cell.value, str):
        data_cell = cell.value
    else:
        data_cell = "="
        page_title = cell.parent.title
        for address in source_files:
            if data_cell[-1] != "=":
                data_cell += "+"
            data_cell += "\'"
            data_cell += os.path.dirname(address)
            data_cell += os.sep
            data_cell += "["
            data_cell += os.path.basename(address)
            data_cell += "]"
            data_cell += page_title
            data_cell += "\'"
            data_cell += "!"
            data_cell += cell.coordinate
    return data_cell


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
    # 2. Отфильтровываем неподходящие (есть нужные страницы или нет, большой файл или нет)
    filteredExcelFileList = check_files(file_list, filter_have_sheet)
    logs.write(f"\nБудут использоваться файлы {filteredExcelFileList}")
    print(filteredExcelFileList)
    # 2.5
    normalize(filteredExcelFileList)
    # 3.
    dir_list = get_directory_sorted_list(filteredExcelFileList)
    example_file = openpyxl.load_workbook(filename="example.xlsx", read_only=True)
    example_file.close()
    sheet_names = example_file.sheetnames
    created_file = file_generator(dir_list)
    for workbook in created_file:
        create_needed_sheet(workbook,sheet_names)
        directory_needed_files = get_current_file_from_directory(filteredExcelFileList, os.path.abspath(workbook[:-5]))
        create_sheet_content(workbook, directory_needed_files)

    logs.close # close logs file
    #abra = input("press any button")
    end = time.time()

    print(end - start)

if __name__ == "__main__":
    main()
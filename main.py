#
# Принцип работы программы:
#   1. смотрим какие есть папки с файлами, создаём списки, вложенные и передаём дальше
#   2. проверяем файлы в папке на то устраивают они нас или нет, выкидываем ненужные
#   3. создаём файл  консолидирующие каждую папку
#   4. заполняем файл данными, формулами, либо обычной суммой, 
#   5. .... далее пока дожить надо
#
import os, openpyxl

logs = open("logs.txt", "a")
logs.truncate(0)

# 1.
def getFileList(dir): # отдаю функции список файлов и папок
    fileList = list() # создаю пустой список
    if not len(dir): # если полученный ранее список пустой
        return fileList # то возвращаю созданный мной пустой список
    else: # если список не пустые
        if os.path.isfile(dir[0]) and dir[0].endswith('.xlsx'): # то беру первый элимент и проверяю excel файл он?
            fileList.append(dir.pop(0)) # если так то удаю из полученного списка и вношу в пустой
        elif os.path.isdir(dir[0]): # если это не excel файл
            getFileList(os.listdir(dir.pop(0))) # то вызываю себяже и отдаю себе перый элемент списка
        else:
            dir.pop(0) # удаляю элемент
            getFileList(dir) # вызываю функцию без удалённого ранее элемента


directorylist = getFileList(os.listdir())

#directory = [item for item in os.listdir() if os.path.isdir(item) and len([file for file in os.listdir(os.path.join(os.getcwd(),item) if file.endwith('.xlsx'))])]
#item.endswith('.xlsx')
#logs.write(f"В работу попадают следующие папки {directory}")
# 2.

# 3.


logs.close # close logs file
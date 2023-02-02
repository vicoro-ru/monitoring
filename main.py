import os, openpyxl


logs = open("logs.txt", "a") # open/create logs file
logs.truncate(0) # clean logs file

#get folders list that have xlsx files
directory = [item for item in os.listdir() if os.path.isdir(item) and not item.endswith('.xlsx')]

logs.write(str(directory))

logs.close # close logs file
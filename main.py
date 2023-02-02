import os, openpyxl

directory = [item for item in os.listdir() if os.path.isdir(item) and True]

print(directory)
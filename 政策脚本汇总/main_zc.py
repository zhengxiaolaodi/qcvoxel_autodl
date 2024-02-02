import xlrd
import pandas as pd
import os
import xlwt
import time
from datetime import datetime
import os
import numpy as np

def get_files_in_folder(folder_path):
    file_list = []
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        if os.path.isfile(file_path):
            file_list.append(file_path)
    return file_list


def xlsx_remake(file_name):
    readfile = xlrd.open_workbook(file_name)
    names = readfile.sheet_names()
    workbook = xlwt.Workbook()
    for i in range(len(names)):
        obj_sheet = readfile.sheet_by_name(names[i])
        row = obj_sheet.nrows
        col = obj_sheet.ncols
        sheet = workbook.add_sheet(names[i])
        for j in range(col):
            result=obj_sheet.cell_value(1, j)
            for i in range(2,row):
                if obj_sheet.cell_value(i, j)!="":
                   result=obj_sheet.cell_value(i, j)
                   sheet.write(i, j, result)
                else: 
                   sheet.write(i, j, result)
    workbook.save("新"+file_name)


file_name="各地市政策信息汇编(2023) .xls"
readfile = xlrd.open_workbook(file_name)
xlsx_remake(file_name)


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
folder_path="./xls"
result_list=[]
file_list=get_files_in_folder(folder_path)
print(file_list)
for i in range(len(file_list)):
        tmp=[]
        sheet_name=file_list[i].replace("./xls/新","").replace(".xls","")
        readfile = xlrd.open_workbook(file_list[i])
        obj_sheet = readfile.sheet_by_name(sheet_name)
        row = obj_sheet.nrows
        col = obj_sheet.ncols
        zc_num=2
        if sheet_name=="江苏省":
           zc_num=1
        result=obj_sheet.cell_value(2,zc_num)
        tmp.append(sheet_name)
        tmp.append(result)
        for i in range(2,row):
            if result!=obj_sheet.cell_value(i,zc_num):
               result=obj_sheet.cell_value(i,zc_num)
               tmp.append(result)
        result_list.append(tmp)
result_np=np.zeros([15,84])
workbook = xlwt.Workbook()
sheet = workbook.add_sheet('Sheet1')
for i in range(len(result_list)):
    for j in range(len(result_list[i])):
        sheet.write(j, i, result_list[i][j])
workbook.save('test.xls')


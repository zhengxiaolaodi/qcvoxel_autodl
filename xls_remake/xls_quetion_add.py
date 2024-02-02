import xlrd

import pandas as pd
import os
import xlwt
from datetime import datetime

def check_project_number(start,end,quary):
    project_number=obj_sheet.cell_value(start, 2)
    for i in range(start,end+1):
        
    




def question_add():
    obj_sheet = readfile.sheet_by_name(names[i])
    row = obj_sheet.nrows
    col = obj_sheet.ncols
    policy_name=obj_sheet.cell_value(2, 3)
    for i in (2,row):
        if obj_sheet.cell_value(i, 3)!=policy_name:
           policy_name=obj_sheet.cell_value(i, 3)
           quary.append(policy_name+"政策文件中，包含了哪些申报项目，")
           
        else:
           answer.append()
           



readfile = xlrd.open_workbook("./各地市政策信息汇编(2023) .xls")
names = readfile.sheet_names()
for i in range(len(names)):
    print("这是"+names[i]+"的政策")
    obj_sheet = readfile.sheet_by_name(names[i])
    row = obj_sheet.nrows
    col = obj_sheet.ncols
    a=[]
    b=[]
    for j in range(col):
         a.append(obj_sheet.cell_value(1, j))
         b.append(j)
    dictionary = dict(zip(a, b))
    print(dictionary)
file_name="文件名"
porject="申报项目"
time_name="发布时间"
content="具体内容"
city="地市"

city_list1=["无锡市"，"南通市"，"镇江市"，"扬州市"，"泰州市"，"徐州市"，"宿迁市"]
city_nanjing="南京市"
city_suzhou="苏州市"
city_huaian="淮安市"
city_changzhou="常州市"
city_yangcheng="盐城市"
city_liangyungang="连云港市"

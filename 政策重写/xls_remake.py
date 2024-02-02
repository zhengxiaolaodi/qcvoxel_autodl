import xlrd

import pandas as pd
import os
import xlwt
from datetime import datetime
def data_change(para):
    delta = pd.Timedelta(str(para)+'days')
    time = pd.to_datetime('1899-12-10')+delta
    return time
def ConvertDate(xlDate):
    # 这里delta是两个日期间隔天数的意思，取值就是Excel来的数字
    delta=datetime.timedelta(days=xlDate)
    # 基础日期就是1899-12-30，将间隔天数加到这个日期上，得到正确的日期戳
    today=datetime.datetime.strptime('1899-12-30','%Y-%m-%d')+delta
    # 格式化输出正确的日期
    return datetime.datetime.strftime(today,'%Y-%m-%d')



def questiong_make(read_sheet_name,xlwt_name):

    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet(read_sheet_name)
    obj_sheet = readfile.sheet_by_name(read_sheet_name)
    row = obj_sheet.nrows
    col = obj_sheet.ncols
    a=[]
    b=[]
    for i in range(1,row):
        for j in range(col):
            sheet.write(i, j, obj_sheet.cell_value(i, j))
    workbook.save("新政策/"+xlwt_name)
def xlsx_remake(xlwt_name,sheet_name):
    readfile = xlrd.open_workbook("新政策/"+xlwt_name)
    obj_sheet = readfile.sheet_by_name(sheet_name)
    row = obj_sheet.nrows
    col = obj_sheet.ncols
    a=[]
    b=[]
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet(sheet_name)
    for j in range(col):
        result=obj_sheet.cell_value(1, j)
        for i in range(2,row):
            if obj_sheet.cell_value(i, j)!="":
               result=obj_sheet.cell_value(i, j)
               sheet.write(i, j, result)
            else: 
               sheet.write(i, j, result)
    workbook.save("新政策/"+"新"+xlwt_name)

readfile = xlrd.open_workbook("./各地市政策信息汇编(2023) .xls")
names = readfile.sheet_names()
for i in range(len(names)):
    xlwt_name=names[i]+".xls"
    print("这是"+names[i]+"的政策")
    questiong_make(names[i],xlwt_name)
    xlsx_remake(xlwt_name,names[i])
    

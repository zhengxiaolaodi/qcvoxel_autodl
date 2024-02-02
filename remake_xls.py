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



def questiong_make_5(sheet_name,xlwt_name):

    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet(sheet_name)
    obj_sheet = readfile.sheet_by_name(sheet_name)
    row = obj_sheet.nrows
    col = obj_sheet.ncols
    a=[]
    b=[]
    for i in range(2,row):
        for j in range(col):
            #print("这是第%d行，第%d列"%(i,j))
            #print(obj_sheet.cell_value(i, j))
            sheet.write(i, j, obj_sheet.cell_value(i, j))
    workbook.save(xlwt_name)
def xlsx_remake(xlwt_name):
    readfile = xlrd.open_workbook(xlwt_name)

    obj_sheet = readfile.sheet_by_name("无锡市")
    row = obj_sheet.nrows
    col = obj_sheet.ncols
    a=[]
    b=[]
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("无锡")
    for j in range(col):
        result=obj_sheet.cell_value(2, j)
        for i in range(3,row):
            if obj_sheet.cell_value(i, j)!="":
               result=obj_sheet.cell_value(i, j)
               sheet.write(i, j, result)
               print(result)
            else: 
               sheet.write(i, j, result)
               print(result)
    workbook.save("新无锡市.xls")
            
           
    
xlsx_remake("无锡市.xls")
            
#readfile = xlrd.open_workbook("./各地市政策信息汇编(2023) .xls")

#names = readfile.sheet_names()

#for i in range(4):
#    xlwt_name=names[i]+".xls"
#    print("这是"+names[i]+"的政策")
#    questiong_make_5(names[i],xlwt_name)
    




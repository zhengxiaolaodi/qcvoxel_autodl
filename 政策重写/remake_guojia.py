import xlrd
import pandas as pd
import os
import xlwt
import time
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
def time_check(answor):
    if type(answor)==str:
       answor=answor.replace("\n","")
    else:
       answor=int(answor)
       if answor>40000 and answor<50000:
          dt = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + answor - 2)
          tt = dt.timetuple()
          str_year=str(tt.tm_year)
          str_mon=str(tt.tm_mon)
          str_day=str(tt.tm_mday)
          answor=str_year+"年"+str_mon+"月"+str_day+"日"
       else:
          answor=str(answor)
    return answor


def xlsx_remake(xlwt_name,sheet_name):
    readfile = xlrd.open_workbook(xlwt_name)

    obj_sheet = readfile.sheet_by_name(sheet_name)
    row = obj_sheet.nrows
    col = obj_sheet.ncols
    a=[]
    b=[]
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet(sheet_name)
    for j in range(col):
        result=obj_sheet.cell_value(2, j)
        for i in range(3,row):
            if obj_sheet.cell_value(i, j)!="":
               result=obj_sheet.cell_value(i, j)
               sheet.write(i, j, result)
            else: 
               sheet.write(i, j, result)
    workbook.save("新"+xlwt_name)

def xlsx_remake_5(xlwt_name,str_list):
    readfile = xlrd.open_workbook(xlwt_name)
    obj_sheet = readfile.sheet_by_name(sheet_name)
    row = obj_sheet.nrows
    col = obj_sheet.ncols
    a=[]
    b=[]
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet(sheet_name)
    result=obj_sheet.cell_value(2,3)
    result_list=[]
    for i in range(2,row):
        if result!=obj_sheet.cell_value(i,3):
           result_list.append(i)
    result_list.append(row)
    for j in range(len(result_list)):
        a.append("的名为"+obj_sheet.cell_value(result_list[j],3)+"的支持政策。")
        answer="其发布时间是"+time_check(obj_sheet.cell_value(result_list[j],4))
        for k in range(result_list[j],result_list[j]+1):
            if obj_sheet.cell_value(k,5)!="" and obj_sheet.cell_value(k,5)!="——":
               answer+="其"+str_list[5]+"是"+obj_sheet.cell_value(k,5) 
        b.append(answer)
    dataframe = pd.DataFrame({'Quary':a,'Answer':b})
    dataframe.to_csv(csv_name,index=False,sep=',')

def xlsx_remake_6(xlwt_name,str_list):
    readfile = xlrd.open_workbook(xlwt_name)
    obj_sheet = readfile.sheet_by_name(sheet_name)
    row = obj_sheet.nrows
    col = obj_sheet.ncols
    a=[]
    b=[]
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet(sheet_name)
    result=obj_sheet.cell_value(2,3)
    result_list=[]
    for i in range(2,row):
        if result!=obj_sheet.cell_value(i,3):
           result_list.append(i)
    result_list.append(row)
    for j in range(len(result_list)):
        a.append("的名为"+obj_sheet.cell_value(result_list[j],3)+"的支持政策。")
        answer="其发布时间是"+time_check(obj_sheet.cell_value(result_list[j],4))
        for k in range(result_list[j],result_list[j]+1):
            if obj_sheet.cell_value(k,5)!="" and obj_sheet.cell_value(k,5)!="——":
               answer+="其"+str_list[5]+"是"+obj_sheet.cell_value(k,5)
            if obj_sheet.cell_value(k,6)!="" and obj_sheet.cell_value(k,6)!="——":
               answer+="其"+str_list[6]+"是"+obj_sheet.cell_value(k,6) 
        b.append(answer)
    dataframe = pd.DataFrame({'Quary':a,'Answer':b})
    dataframe.to_csv(csv_name,index=False,sep=',')

def xlsx_remake_7(xlwt_name,str_list):
    readfile = xlrd.open_workbook(xlwt_name)
    obj_sheet = readfile.sheet_by_name(sheet_name)
    row = obj_sheet.nrows
    col = obj_sheet.ncols
    a=[]
    b=[]
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet(sheet_name)
    result=obj_sheet.cell_value(2,3)
    result_list=[]
    for i in range(2,row):
        if result!=obj_sheet.cell_value(i,3):
           result_list.append(i)
    result_list.append(row)
    for j in range(len(result_list)):
        a.append("的名为"+obj_sheet.cell_value(result_list[j],3)+"的支持政策。")
        answer="其发布时间是"+time_check(obj_sheet.cell_value(result_list[j],4))
        for k in range(result_list[j],result_list[j]+1):
            if obj_sheet.cell_value(k,5)!="" and obj_sheet.cell_value(k,5)!="——":
               answer+="其"+str_list[5]+"是"+obj_sheet.cell_value(k,5)
            if obj_sheet.cell_value(k,6)!="" and obj_sheet.cell_value(k,6)!="——":
               answer+="其"+str_list[6]+"是"+obj_sheet.cell_value(k,6)
            if obj_sheet.cell_value(k,7)!="" and obj_sheet.cell_value(k,7)!="——":
               answer+="其"+str_list[7]+"是"+obj_sheet.cell_value(k,7)   
        b.append(answer)
    dataframe = pd.DataFrame({'Quary':a,'Answer':b})
    dataframe.to_csv(csv_name,index=False,sep=',')  



readfile = xlrd.open_workbook("./各地市政策信息汇编(2023) .xlsx")
str_nanjing=[",",
"项目中,",
        ",",
    "的支持政策的名称是",
    "的支持政策发布的时间是",
    "的支持政策的具体内容是",
    "的支持政策的奖补情况是"]

str_suzhou=[",",
"项目中,",
        ",",
    "的支持政策的名称是",
    "的支持政策发布的时间是",
    "的支持政策的申报时间和类别是",
    "的支持政策规定的申报方式是",
    "的支持政策的具体内容是"]
str_huaian=[",",
"项目中,",
        ",",
    "的支持政策的名称是",
    "的支持政策发布的时间是",
    "的支持政策的申报方式是",
    "的支持政策的奖补情况是"]
str_changzhou=[",",
"项目中,",
        ",",
    "的支持政策的名称是",
    "的支持政策发布的时间是",
    "的支持政策的申报时间和方式是",
    "的支持政策发布的第二个文件名是",
    "的支持政策发布的第二个文件的申报内容是"]
str_yanchen=[",",
            ",",
    "项目中,",
           ",",
    "的支持政策发布的时间是",
    "的支持政策的名称是",
    "的支持政策的具体内容是"]
str_lianyungang=[",",
"项目中,",
        ",",
    "的支持政策的名称是",
    "的支持政策发布的时间是",
    "的支持政策的具体内容是",
    "的支持政策的类别是",
    "的支持政策的奖励金额是"]
str_else=[",",
"项目中,",
        ",",
    "的支持政策的名称是",
    "的支持政策发布的时间是",
    "的支持政策的具体内容是",
    ]

    
names = readfile.sheet_names()
for i in range(len(names)):
    csv_name="政策.csv"
    csv_name=names[i]+csv_name
    if names[i] == "南京市":
       xlsx_remake_6(csv_name,str_nanjing)
    elif names[i] == "淮安市":
       xlsx_remake_6(csv_name,str_huaian)
    elif names[i] == "苏州市":
       xlsx_remake_7(csv_name,str_suzhou)
    elif names[i] == "常州市":
       xlsx_remake_7(csv_name,str_changzhou)
    elif names[i] == "盐城市":
       questiong_make_yanc(csv_name,str_yanchen)
    elif names[i] == "连云港市":
       xlsx_remake_7(csv_namestr_lianyungang)
    else:
       xlsx_remake_5(csv_name,str_else)
    

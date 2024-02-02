import xlrd
import pandas as pd
import os
import xlwt
import time
from datetime import datetime
import os
 
def get_files_in_folder(folder_path):
    file_list = []
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        if os.path.isfile(file_path):
            file_list.append(file_path)
    return file_list
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


def xlsx_remake(xlwt_name,sheet_name,csv_name):
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

def xlsx_remake_5(xlwt_name,sheet_name,str_list,csv_name):
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
    result_list.append(2)
    answer=result
    for i in range(2,row):
        if result!=obj_sheet.cell_value(i,3):
           result_list.append(i)
           print(result)
           result=obj_sheet.cell_value(i,3)
    result_list.append(row)
    print(result_list)
    for j in range(len(result_list)-1):
        a.append(obj_sheet.cell_value(result_list[j],1)+"的名为"+(obj_sheet.cell_value(result_list[j],3)+"的政策").replace("\n","").replace(" ",""))
        answer="支持政策发布时间是"+time_check(obj_sheet.cell_value(result_list[j],4))+"。"
        for k in range(result_list[j],result_list[j+1]):
            if obj_sheet.cell_value(k,5)!="" and obj_sheet.cell_value(k,5)!="——":
               answer+="其"+str_list[5]+obj_sheet.cell_value(k,5)+"。"
        b.append(answer.replace("\n","").replace(" ",""))
    dataframe = pd.DataFrame({'Quary':a,'Answer':b})
    dataframe.to_csv("政策"+csv_name,index=False,sep=',')

def xlsx_remake_6(xlwt_name,sheet_name,str_list,csv_name):
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
    result_list.append(2)
    answer=result
    for i in range(2,row):
        if result!=obj_sheet.cell_value(i,3):
           result_list.append(i)
           print(result)
           result=obj_sheet.cell_value(i,3)
    result_list.append(row)
    for j in range(len(result_list)-1):
        a.append(obj_sheet.cell_value(result_list[j],1)+"的名为"+(obj_sheet.cell_value(result_list[j],3)+"的政策").replace("\n","").replace(" ",""))
        answer="支持政策发布时间是"+time_check(obj_sheet.cell_value(result_list[j],4))+"。"
        for k in range(result_list[j],result_list[j+1]):      
            if obj_sheet.cell_value(k,5)!="" and obj_sheet.cell_value(k,5)!="——":
               answer+="其"+str_list[5]+obj_sheet.cell_value(k,5)+";"
            if obj_sheet.cell_value(k,6)!="" and obj_sheet.cell_value(k,6)!="——":
               answer+="其"+str_list[6]+obj_sheet.cell_value(k,6) +"。"  
        b.append(answer.replace("\n","").replace(" ",""))
    dataframe = pd.DataFrame({'Quary':a,'Answer':b})
    dataframe.to_csv("政策"+csv_name,index=False,sep=',')


def xlsx_remake_guojia(xlwt_name,sheet_name,str_list,csv_name):
    readfile = xlrd.open_workbook(xlwt_name)
    obj_sheet = readfile.sheet_by_name(sheet_name)
    row = obj_sheet.nrows
    col = obj_sheet.ncols
    a=[]
    b=[]
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet(sheet_name)
    result=obj_sheet.cell_value(2,2)
    result_list=[]
    result_list.append(2)
    answer=result
    for i in range(2,row):
        if result!=obj_sheet.cell_value(i,2):
           result_list.append(i)
           print(result)
           result=obj_sheet.cell_value(i,2)
    result_list.append(row)
    for j in range(len(result_list)-1):
        a.append(obj_sheet.cell_value(result_list[j],1)+"的名为"+(obj_sheet.cell_value(result_list[j],3)+"的政策").replace("\n","").replace(" ",""))
        answer="支持政策发布时间是"+time_check(obj_sheet.cell_value(result_list[j],4))+"。"
        for k in range(result_list[j],result_list[j+1]):
            if obj_sheet.cell_value(k,5)!="" and obj_sheet.cell_value(k,5)!="——":
               answer+="其"+str_list[5]+obj_sheet.cell_value(k,5)+";"
            if obj_sheet.cell_value(k,6)!="" and obj_sheet.cell_value(k,6)!="——":
               answer+="其"+str_list[6]+obj_sheet.cell_value(k,6)+"。"  
        b.append(answer.replace("\n","").replace(" ",""))
    dataframe = pd.DataFrame({'Quary':a,'Answer':b})
    dataframe.to_csv("政策"+csv_name,index=False,sep=',')




def xlsx_remake_yanc(xlwt_name,sheet_name,str_list,csv_name):
    readfile = xlrd.open_workbook(xlwt_name)
    obj_sheet = readfile.sheet_by_name(sheet_name)
    row = obj_sheet.nrows
    col = obj_sheet.ncols
    a=[]
    b=[]
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet(sheet_name)
    result=obj_sheet.cell_value(2,5)
    result_list=[]
    result_list.append(2)
    answer=result
    for i in range(2,row):
        if result!=obj_sheet.cell_value(i,5):
           result_list.append(i)
    result_list.append(row)
    for j in range(len(result_list)-1):
        a.append((obj_sheet.cell_value(result_list[j],5)+"的项目。").replace("\n","").replace(" ",""))
        answer="支持政策发布时间是"+time_check(obj_sheet.cell_value(result_list[j],4))+"。"
        for k in range(result_list[j],result_list[j+1]):
            if obj_sheet.cell_value(k,2)!="" and obj_sheet.cell_value(k,2)!="——":
               answer+="其"+str_list[2]+obj_sheet.cell_value(k,2)+";"
            if obj_sheet.cell_value(k,6)!="" and obj_sheet.cell_value(k,6)!="——":
               answer+="其"+str_list[6]+obj_sheet.cell_value(k,6) +"。"
        b.append(answer.replace("\n","").replace(" ",""))
    dataframe = pd.DataFrame({'Quary':a,'Answer':b})
    dataframe.to_csv("政策"+csv_name,index=False,sep=',')

def xlsx_remake_jiangsusheng(xlwt_name,sheet_name,str_list,csv_name):
    readfile = xlrd.open_workbook(xlwt_name)
    obj_sheet = readfile.sheet_by_name(sheet_name)
    row = obj_sheet.nrows
    col = obj_sheet.ncols
    a=[]
    b=[]
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet(sheet_name)
    result=obj_sheet.cell_value(2,2)
    result_list=[]
    result_list.append(2)
    a.append(sheet_name+"的支持政策文件有哪些？")
    answer=result
    for i in range(2,row):
        if result!=obj_sheet.cell_value(i,2):
           result_list.append(i)
           print(result)
           result=obj_sheet.cell_value(i,2)
           answer+=result+"。"
    b.append(answer.replace("\n","").replace(" ",""))
    result_list.append(row)
    for j in range(len(result_list)-1):
        a.append(("江苏省的名为"+obj_sheet.cell_value(result_list[j],2)+"的政策").replace("\n","").replace(" ",""))
        answer="政策发布的时间是"+time_check(obj_sheet.cell_value(result_list[j],3))+"。"
        for k in range(result_list[j],result_list[j+1]):
            if obj_sheet.cell_value(k,4)!="" and obj_sheet.cell_value(k,4)!="——":
               answer+="其"+str_list[4]+obj_sheet.cell_value(k,4)+";"
            if obj_sheet.cell_value(k,5)!="" and obj_sheet.cell_value(k,5)!="——":
               answer+="其"+str_list[5]+obj_sheet.cell_value(k,5)+"。"
        b.append(answer.replace("\n","").replace(" ",""))
    dataframe = pd.DataFrame({'Quary':a,'Answer':b})
    dataframe.to_csv("政策"+csv_name,index=False,sep=',')




def xlsx_remake_7(xlwt_name,sheet_name,str_list,csv_name):
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
    result_list.append(2)
    answer=result
    for i in range(2,row):
        if result!=obj_sheet.cell_value(i,3):
           result_list.append(i)
           print(result)
           result=obj_sheet.cell_value(i,3)
    result_list.append(row)
    for j in range(len(result_list)-1):
        a.append((obj_sheet.cell_value(result_list[j],1)+"的名为"+obj_sheet.cell_value(result_list[j],3)+"的项目。").replace("\n","").replace(" ",""))
        answer=obj_sheet.cell_value(result_list[j],1)+"的支持政策发布时间是"+time_check(obj_sheet.cell_value(result_list[j],4))+"。"
        for k in range(result_list[j],result_list[j+1]):
            if obj_sheet.cell_value(k,5)!="" and obj_sheet.cell_value(k,5)!="——":
               answer+="其"+str_list[5]+obj_sheet.cell_value(k,5)+";"
            if obj_sheet.cell_value(k,6)!="" and obj_sheet.cell_value(k,6)!="——":
               answer+="其"+str_list[6]+obj_sheet.cell_value(k,6)+";"
            if obj_sheet.cell_value(k,7)!="" and obj_sheet.cell_value(k,7)!="——":
               answer+="其"+str_list[7]+obj_sheet.cell_value(k,7)+"。"   
        b.append(answer.replace("\n","").replace(" ",""))
    dataframe = pd.DataFrame({'Quary':a,'Answer':b})
    dataframe.to_csv("政策"+csv_name,index=False,sep=',')  

def xlsx_remake_changzhou(xlwt_name,sheet_name,str_list,csv_name):
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
    result_list.append(2)
    answer=result
    for i in range(2,row):
        if result!=obj_sheet.cell_value(i,3):
           result_list.append(i)
           print(result)
           result=obj_sheet.cell_value(i,3)
    result_list.append(row)
    for j in range(len(result_list)-1):
        a.append((obj_sheet.cell_value(result_list[j],1)+"的名为"+obj_sheet.cell_value(result_list[j],3)+"的政策。").replace("\n","").replace(" ",""))
        answer=obj_sheet.cell_value(result_list[j],1)+"的支持政策的发布时间是"+time_check(obj_sheet.cell_value(result_list[j],4))+"。"
        for k in range(result_list[j],result_list[j+1]):
            if obj_sheet.cell_value(k,5)!="" and obj_sheet.cell_value(k,5)!="——":
               answer+="其"+str_list[5]+obj_sheet.cell_value(k,5)+"。"
    
        b.append(answer.replace("\n","").replace(" ",""))
    dataframe = pd.DataFrame({'Quary':a,'Answer':b})
    dataframe.to_csv("政策"+csv_name,index=False,sep=',')  



str_guojia=["序号",
"所属类别,",
        "项目名词",
    "支持政策的文件是",
    "发布时间是",
    "具体内容是",
    "目标是"]
str_jiangsus=["序号",
"申报项目",
        "文件名",
    "发布时间",
    "项目将补申请是",
    "具体内容是"]



str_nanjing=[",",
"项目中,",
        ",",
    "支持政策的名称是",
    "支持政策发布的时间是",
    "支持政策的具体内容是",
    "支持政策的奖补情况是"]

str_suzhou=[",",
"项目中,",
        ",",
    "支持政策的名称是",
    "支持政策发布的时间是",
    "支持政策的申报时间和类别是",
    "支持政策规定的申报方式是",
    "支持政策的具体内容是"]
str_huaian=[",",
"项目中,",
        ",",
    "支持政策的名称是",
    "支持政策发布的时间是",
    "支持政策的申报方式是",
    "支持政策的奖补情况是"]
str_changzhou=[",",
"项目中,",
        ",",
    "支持政策的名称是",
    "支持政策发布的时间是",
    "支持政策的申报时间和方式是",
    "支持政策发布的第二个文件名是",
    "支持政策发布的第二个文件的申报内容是"]
str_yanchen=[",",
            ",",
    "项目中,",
           ",",
    "支持政策发布的时间是",
    "支持政策的名称是",
    "支持政策的具体内容是"]
str_lianyungang=[",",
"项目中,",
        ",",
    "支持政策的名称是",
    "支持政策发布的时间是",
    "支持政策的具体内容是",
    "支持政策的类别是",
    "支持政策的奖励金额是"]
str_else=[",",
"项目中,",
        ",",
    "支持政策的名称是",
    "支持政策发布的时间是",
    "支持政策的具体内容是",
    ]
    

    
xin="新"
#print(file_list[i])
folder_path="./xls"
file_list=get_files_in_folder(folder_path)
for i in range(len(file_list)):
    sheet_name=file_list[i].replace("新", "").replace("./xls/","").replace(".xls","")
    csv_name="政策.csv"
    csv_name=sheet_name.replace(".xls","")+csv_name
    if sheet_name == "南京市":
       xlsx_remake_6(file_list[i],sheet_name,str_nanjing,csv_name)
       print("sheet_name")
       print(sheet_name)
    elif sheet_name == "淮安市":
       xlsx_remake_6(file_list[i],sheet_name,str_huaian,csv_name)
       print("sheet_name")
       print(sheet_name)
    elif sheet_name == "苏州市":
       xlsx_remake_7(file_list[i],sheet_name,str_suzhou,csv_name)
       print("sheet_name")
       print(sheet_name)
    elif sheet_name == "常州市":
       xlsx_remake_changzhou(file_list[i],sheet_name,str_changzhou,csv_name)
       print("sheet_name")
       print(sheet_name)
    elif sheet_name == "盐城市":
       xlsx_remake_yanc(file_list[i],sheet_name,str_yanchen,csv_name)
       print("file_list")
       print(file_list[i])
    elif sheet_name == "连云港市":
       xlsx_remake_7(file_list[i],sheet_name,str_lianyungang,csv_name)
       print("sheet_name")
       print(sheet_name)
    elif sheet_name == "国家":
       xlsx_remake_guojia(file_list[i],sheet_name,str_guojia,csv_name)
       print("sheet_name")
       print(sheet_name)
    elif sheet_name == "江苏省":
       xlsx_remake_jiangsusheng(file_list[i],sheet_name,str_jiangsus,csv_name)
       print("sheet_name")
       print(sheet_name)
    else:
       xlsx_remake_5(file_list[i],sheet_name,str_else,csv_name)
       print("sheet_name")
       print(sheet_name)

    

import xlrd

import pandas as pd
import os

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



def questiong_make_5(csv_name,sheet_name,str_city):
    obj_sheet = readfile.sheet_by_name(sheet_name)
    row = obj_sheet.nrows
    col = obj_sheet.ncols
    a=[]
    b=[]
    for i in range(2,row):
        if obj_sheet.cell_value(i, 1)!="" and obj_sheet.cell_value(i, 1)!="——":
           main_name=obj_sheet.cell_value(i, 1)
        else:
           main_name=sheet_name
        if obj_sheet.cell_value(i, 3)!=""and obj_sheet.cell_value(i, 3)!="——":
           a.append(obj_sheet.cell_value(i, 2)+str_city[1]+main_name+str_city[3])
           answor=obj_sheet.cell_value(i, 3)
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
           b.append(answor) 
        if obj_sheet.cell_value(i, 4)!=""and obj_sheet.cell_value(i, 4)!="——":
           a.append(obj_sheet.cell_value(i, 2)+str_city[1]+main_name+str_city[4])
           answor=obj_sheet.cell_value(i, 4)
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
           b.append(answor)    
        if obj_sheet.cell_value(i, 5)!=""and obj_sheet.cell_value(i, 5)!="——":
           a.append(obj_sheet.cell_value(i, 2)+str_city[1]+main_name+str_city[5])
           answor=obj_sheet.cell_value(i, 5)
           answor=answor.replace("\n","")
           b.append(answor)
    dataframe = pd.DataFrame({'Quary':a,'Answer':b})
    dataframe.to_csv(csv_name,index=False,sep=',')  

def questiong_make_6(csv_name,sheet_name,str_city):
    obj_sheet = readfile.sheet_by_name(sheet_name)
    row = obj_sheet.nrows
    col = obj_sheet.ncols
    a=[]
    b=[]
    for i in range(2,row):
        if obj_sheet.cell_value(i, 1)!="" and obj_sheet.cell_value(i, 1)!="——":
           main_name=obj_sheet.cell_value(i, 1)
        else:
           main_name=sheet_name
        if obj_sheet.cell_value(i, 3)!="" and obj_sheet.cell_value(i, 3)!="——":
           a.append(obj_sheet.cell_value(i, 2)+str_city[1]+main_name+str_city[3])
           answor=obj_sheet.cell_value(i, 3)
           answor=answor.replace("\n","")
           b.append(answor)
        if obj_sheet.cell_value(i, 4)!="" and obj_sheet.cell_value(i, 3)!="——":
           a.append(obj_sheet.cell_value(i, 2)+str_city[1]+main_name+str_city[4])
           answor=obj_sheet.cell_value(i, 4)
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
           b.append(answor)    
        if obj_sheet.cell_value(i, 5)!="" and obj_sheet.cell_value(i, 5)!="——":
           a.append(obj_sheet.cell_value(i, 2)+str_city[1]+main_name+str_city[5])
           answor=obj_sheet.cell_value(i, 5)
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
           b.append(answor)
        if obj_sheet.cell_value(i, 6)!="" and obj_sheet.cell_value(i, 6)!="——":
           a.append(obj_sheet.cell_value(i, 2)+str_city[1]+main_name+str_city[6])
           answor=obj_sheet.cell_value(i, 6)
           answor=answor.replace("\n","")
           b.append(answor)
    dataframe = pd.DataFrame({'Quary':a,'Answer':b})
    dataframe.to_csv(csv_name,index=False,sep=',')    


def questiong_make_yanc(csv_name,sheet_name,str_city):
    obj_sheet = readfile.sheet_by_name(sheet_name)
    row = obj_sheet.nrows
    col = obj_sheet.ncols
    a=[]
    b=[]
    for i in range(2,row):
        if obj_sheet.cell_value(i, 1)!="" and obj_sheet.cell_value(i, 1)!="——":
           main_name=obj_sheet.cell_value(i, 1)
           print(i)
        else:
           main_name=sheet_name  
        print(main_name)
        if obj_sheet.cell_value(i, 4)!="" and obj_sheet.cell_value(i, 4)!="——":
           a.append(obj_sheet.cell_value(i, 2)+str_city[1]+main_name+str_city[4])
           answor=obj_sheet.cell_value(i, 4)
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
           b.append(answor)    
        if obj_sheet.cell_value(i, 5)!="" and obj_sheet.cell_value(i, 5)!="——":
           a.append(obj_sheet.cell_value(i, 2)+str_city[1]+main_name+str_city[5])
           answor=obj_sheet.cell_value(i, 5)
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
           b.append(answor)
        if obj_sheet.cell_value(i, 6)!="" and obj_sheet.cell_value(i, 6)!="——":
           a.append(obj_sheet.cell_value(i, 2)+str_city[1]+main_name+str_city[6])
           answor=obj_sheet.cell_value(i, 6)
           answor=answor.replace("\n","")
           b.append(answor)
    dataframe = pd.DataFrame({'Quary':a,'Answer':b})
    dataframe.to_csv(csv_name,index=False,sep=',')    


def questiong_make_7(csv_name,sheet_name,str_city,start=2):
    obj_sheet = readfile.sheet_by_name(sheet_name)
    row = obj_sheet.nrows
    col = obj_sheet.ncols
    a=[]
    b=[]
    for i in range(start,row):
        if obj_sheet.cell_value(i, 1)!="" and obj_sheet.cell_value(i, 1)!="——":
           main_name=obj_sheet.cell_value(i, 1)
        else:
           main_name=sheet_name
        if obj_sheet.cell_value(i, 3)!="" and obj_sheet.cell_value(i, 3)!="——":
           a.append(obj_sheet.cell_value(i, 2)+str_city[1]+main_name+str_city[3])
           answor=obj_sheet.cell_value(i, 3)
           answor=answor.replace("\n","")
           b.append(answor)
        if obj_sheet.cell_value(i, 4)!="" and obj_sheet.cell_value(i, 3)!="——":
           a.append(obj_sheet.cell_value(i, 2)+str_city[1]+main_name+str_city[4])
           answor=obj_sheet.cell_value(i, 4)
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
           b.append(answor)    
        if obj_sheet.cell_value(i, 5)!="" and obj_sheet.cell_value(i, 5)!="——":
           a.append(obj_sheet.cell_value(i, 2)+str_city[1]+main_name+str_city[5])
           answor=obj_sheet.cell_value(i, 5)
           answor=answor.replace("\n","")
           b.append(answor)
        if obj_sheet.cell_value(i, 6)!="" and obj_sheet.cell_value(i, 6)!="——":
           a.append(obj_sheet.cell_value(i, 2)+str_city[1]+main_name+str_city[6])
           answor=obj_sheet.cell_value(i, 6)
           answor=answor.replace("\n","")
           b.append(answor)
        if obj_sheet.cell_value(i, 6)!="" and obj_sheet.cell_value(i, 6)!="——":
           a.append(obj_sheet.cell_value(i, 2)+str_city[1]+main_name+str_city[6])
           answor=obj_sheet.cell_value(i, 6)
           answor=answor.replace("\n","")
           b.append(answor)
    dataframe = pd.DataFrame({'Quary':a,'Answer':b})
    dataframe.to_csv(csv_name,index=False,sep=',')  






readfile = xlrd.open_workbook("./各地市政策信息汇编(2023) .xls")
str_yancheng=[",",
"项目中,",
        ",",
    "的支持政策的名称是",
    "的支持政策发布的时间是",
    "的支持政策的具体内容是",
    "的支持政策的奖补情况是"]

names = "盐城市"
csv_name="政策.csv"
csv_name=names+csv_name
questiong_make_6(csv_name,names,str_yancheng)

       

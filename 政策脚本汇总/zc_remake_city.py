import xlrd
import pandas as pd
import os
import xlwt
import time
from datetime import datetime
import os
#这里是问城市有哪些政策



class Zc_Q_City_A_Zcname:  
    # 类的属性  
    str_guojia=["序号",
"所属类别,",
        "项目名词",
    "文件名",
    "发布时间",
    "具体内容",
    "目标"]
    str_jiangsus=["序号",
"申报项目",
        "文件名",
    "发布时间",
    "项目将补申请",
    "具体内容是",
    "链接"]
    str_nanjing=["序号",
    "地市",
        "申报项目",
    "发布时间",
    "具体内容",
    "申报时间和方式",
    "奖补情况",
    "链接"]

    str_suzhou=["序号",
    "地市",
        "申报项目",
    "文件名",
    "发布的时间",
    "申报时间和类别",
    "申报方式",
    "具体内容"，
    "链接"]
    str_huaian=["序号",
      "地市",
        "申报项目",
    "文件名",
    "发布时间",
    "申报方式",
    "奖补情况",
    "链接"]
    str_changzhou=["序号",
    "地市",
        "申报项目",
    "文件名",
    "发布的时间",
    "申报时间和方式",
    "第二个文件名",
    "第二个文件的申报内容",
    "链接"]
    str_yanchen=["序号",
            "地市",
    "申报项目",
           "企业条件",
    "发布时间",
    "文件名",
    "具体内容",
    "链接"]
    str_lianyungang=["序号",
      "地市",
        "申报项目",
    "文件名",
    "发布的时间",
    "具体内容",
    "类别",
    "奖励金额",
    "链接"]
    str_else=["序号",
    "地市",
    "申报项目",
    "文件名",
    "发布的时间",
    "具体内容",
    "链接"
    ]  
    # 类的方法  
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

    def xlsx_remake_5(xlwt_name,sheet_name,str_list,csv_name):
        readfile = xlrd.open_workbook(xlwt_name)
        obj_sheet = readfile.sheet_by_name(sheet_name)
        row = obj_sheet.nrows
        col = obj_sheet.ncols
        a=[]
        b=[]
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet(sheet_name)
        result=obj_sheet.cell_value(2,1)
        result_list=[]
        result_list.append(2)
        for i in range(2,row):
            if result!=obj_sheet.cell_value(i,1):
               result_list.append(i)
               result=obj_sheet.cell_value(i,1)
        result_list.append(row)
        print(result_list)
        for j in range(len(result_list)-1):
            a.append((obj_sheet.cell_value(result_list[j],1)+"的政策有哪些?").replace("\n","").replace(" ",""))
            answer="政策包括："
            for k in range(result_list[j],result_list[j+1]):
                if obj_sheet.cell_value(k,3)!="" and obj_sheet.cell_value(k,3)!="——":
                   answer+=obj_sheet.cell_value(k,3)+"。"
            b.append(answer.replace("\n","").replace(" ",""))
        dataframe = pd.DataFrame({'Quary':a,'Answer':b})
        dataframe.to_csv("城市_"+csv_name,index=False,sep=',')

    def xlsx_remake_6(xlwt_name,sheet_name,str_list,csv_name):
        readfile = xlrd.open_workbook(xlwt_name)
        obj_sheet = readfile.sheet_by_name(sheet_name)
        row = obj_sheet.nrows
        col = obj_sheet.ncols
        a=[]
        b=[]
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet(sheet_name)
        result=obj_sheet.cell_value(2,1)
        result_list=[]
        result_list.append(2)
        for i in range(2,row):
            if result!=obj_sheet.cell_value(i,1):
               result_list.append(i)
               result=obj_sheet.cell_value(i,1)
        result_list.append(row)
        print(result_list)
        for j in range(len(result_list)-1):
            a.append((obj_sheet.cell_value(result_list[j],1)+"的政策有哪些?").replace("\n","").replace(" ",""))
            answer="政策包括："
            for k in range(result_list[j],result_list[j+1]):
                if obj_sheet.cell_value(k,3)!="" and obj_sheet.cell_value(k,3)!="——":
                   answer+=obj_sheet.cell_value(k,3)+"。"
            b.append(answer.replace("\n","").replace(" ",""))
        dataframe = pd.DataFrame({'Quary':a,'Answer':b})
        dataframe.to_csv("城市_"+csv_name,index=False,sep=',')


    def xlsx_remake_guojia(xlwt_name,sheet_name,str_list,csv_name):
        readfile = xlrd.open_workbook(xlwt_name)
        obj_sheet = readfile.sheet_by_name(sheet_name)
        row = obj_sheet.nrows
        col = obj_sheet.ncols
        a=[]
        b=[]
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet(sheet_name)
        result=obj_sheet.cell_value(2,1)
        result_list=[]
        result_list.append(2)
        for i in range(2,row):
            if result!=obj_sheet.cell_value(i,1):
               result_list.append(i)
               result=obj_sheet.cell_value(i,1)
        result_list.append(row)
        for j in range(len(result_list)-1):
            a.append(("国家的类别"+obj_sheet.cell_value(j,1)+"的支持政策包括：").replace("\n","").replace(" ",""))
            answer="政策包括："
            for k in range(result_list[j],result_list[j+1]):
                if obj_sheet.cell_value(k,3)!="" and obj_sheet.cell_value(k,3)!="——":
                   answer+=obj_sheet.cell_value(k,3)+"。"  
            b.append(answer.replace("\n","").replace(" ",""))
        dataframe = pd.DataFrame({'Quary':a,'Answer':b})
        dataframe.to_csv("城市_"+csv_name,index=False,sep=',')




    def xlsx_remake_yanc(xlwt_name,sheet_name,str_list,csv_name):
        readfile = xlrd.open_workbook(xlwt_name)
        obj_sheet = readfile.sheet_by_name(sheet_name)
        row = obj_sheet.nrows
        col = obj_sheet.ncols
        a=[]
        b=[]
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet(sheet_name)
        result=obj_sheet.cell_value(2,1)
        result_list=[]
        result_list.append(2)
        a.append(sheet_name+"的支持政策文件有哪些？")
        answer=result
        for i in range(2,row):
            if result!=obj_sheet.cell_value(i,1):
               result_list.append(i)
               print(result)
               result=obj_sheet.cell_value(i,1)
               answer+=result+"。"
        b.append(answer.replace("\n","").replace(" ",""))
        result_list.append(row)
        for j in range(len(result_list)-1):
            a.append((obj_sheet.cell_value(result_list[j],2)+"的支持政策有？").replace("\n","").replace(" ",""))
            answer="政策包括："
            for k in range(result_list[j],result_list[j+1]):
                if obj_sheet.cell_value(k,5)!="" and obj_sheet.cell_value(k,5)!="——":
                   answer+=obj_sheet.cell_value(k,5)+"。"
            b.append(answer.replace("\n","").replace(" ",""))
        dataframe = pd.DataFrame({'Quary':a,'Answer':b})
        dataframe.to_csv("城市_"+csv_name,index=False,sep=',')

    def xlsx_remake_jiangsusheng(xlwt_name,sheet_name,str_list,csv_name):
        readfile = xlrd.open_workbook(xlwt_name)
        obj_sheet = readfile.sheet_by_name(sheet_name)
        row = obj_sheet.nrows
        col = obj_sheet.ncols
        a=[]
        b=[]
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet(sheet_name)
        result=obj_sheet.cell_value(2,1)
        result_list=[]
        result_list.append(2)
        a.append(sheet_name+"的支持政策文件有哪些？")
        answer=result
        for i in range(2,row):
            if result!=obj_sheet.cell_value(i,1):
               result_list.append(i)
               print(result)
               result=obj_sheet.cell_value(i,1)
               answer+=result+"。"
        b.append(answer.replace("\n","").replace(" ",""))
        result_list.append(row)
        dataframe = pd.DataFrame({'Quary':a,'Answer':b})
        dataframe.to_csv("城市_"+csv_name,index=False,sep=',')




    def xlsx_remake_7(xlwt_name,sheet_name,str_list,csv_name):
        readfile = xlrd.open_workbook(xlwt_name)
        obj_sheet = readfile.sheet_by_name(sheet_name)
        row = obj_sheet.nrows
        col = obj_sheet.ncols
        a=[]
        b=[]
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet(sheet_name)
        result=obj_sheet.cell_value(2,1)
        result_list=[]
        result_list.append(2)
        answer=result
        for i in range(2,row):
            if result!=obj_sheet.cell_value(i,1):
               result_list.append(i)
               result=obj_sheet.cell_value(i,1)
        result_list.append(row)
        for j in range(len(result_list)-1):
            a.append((obj_sheet.cell_value(result_list[j],1)+"的支持政策有？").replace("\n","").replace(" ",""))
            answer="政策包括："
            for k in range(result_list[j],result_list[j+1]):
                if obj_sheet.cell_value(k,3)!="" and obj_sheet.cell_value(k,3)!="——":
                   answer+=obj_sheet.cell_value(k,3)+"。"  
            b.append(answer.replace("\n","").replace(" ",""))
        dataframe = pd.DataFrame({'Quary':a,'Answer':b})
        dataframe.to_csv("城市_"+csv_name,index=False,sep=',')  

    def xlsx_remake_changzhou(xlwt_name,sheet_name,str_list,csv_name):
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
            a.append((obj_sheet.cell_value(result_list[j],1)+"的支持政策有？").replace("\n","").replace(" ",""))
            answer="政策包括："
            for k in range(result_list[j],result_list[j+1]):
                if obj_sheet.cell_value(k,3)!="" and obj_sheet.cell_value(k,3)!="——":
                   answer+=obj_sheet.cell_value(k,3)+"。"      
            b.append(answer.replace("\n","").replace(" ",""))
        dataframe = pd.DataFrame({'Quary':a,'Answer':b})
        dataframe.to_csv("城市_"+csv_name,index=False,sep=',')  








    

    
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

    

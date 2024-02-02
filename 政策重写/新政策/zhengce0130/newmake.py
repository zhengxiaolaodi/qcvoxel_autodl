import xlrd
import pandas as pd
import os
import xlwt
import time
from datetime import datetime
import os
import numpy as np
import re
import xlrd
import jieba.analyse
def remove_urls(text):
    pattern = r'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+'
    result = re.sub(pattern, '', text)
    return result.replace("。。","。").replace("\n","").replace("(suzhou.gov.cn)","").replace("（www.cspiii.com/pg）","").replace("www.dxplat.com","").replace("(changzhou.gov.cn)","").replace("(zjg.gov.cn)","").replace("(tzhl.gov.cn)","").replace("(snd.gov.cn)","").replace("（）","")
def similarity(str1, str2):
    m, n = len(str1), len(str2)
    dp = [[0] * (n+1) for _ in range(m+1)]
 
    for i in range(m+1):
        dp[i][0] = i
    for j in range(n+1):
        dp[0][j] = j
 
    for i in range(1, m+1):
        for j in range(1, n+1):
            if str1[i-1] == str2[j-1]:
                dp[i][j] = dp[i-1][j-1]
            else:
                dp[i][j] = min(dp[i-1][j], dp[i][j-1], dp[i-1][j-1]) + 1
 
    return 1 - dp[m][n] / max(m, n)
    

workbook = xlrd.open_workbook('train.xls')
sheet_names = workbook.sheet_names()
sheet = workbook.sheet_by_index(0)
num_rows = sheet.nrows
num_cols = sheet.ncols
a=[]
b=[]
for i in range(num_rows):
    if sheet.cell_value(i, 0)!="" and sheet.cell_value(i, 1)!="":
       a.append(sheet.cell_value(i, 0))
       b.append(remove_urls(sheet.cell_value(i, 1)))             
dataframe = pd.DataFrame({'Quary':a,'Answer':b})
dataframe.to_csv("新政策.csv",index=False,sep=',')  
result=[]
for m in range(len(a)):
    if len(a[m])<14:
       result.append(a[m].replace("的项目。","").replace("的项目",""))
result = list(set(result))
for j in range(len(result)):
    print(result[j])








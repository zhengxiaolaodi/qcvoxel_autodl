import xlrd
import pandas as pd
import os
import xlwt
import time
from datetime import datetime
import os
import numpy as np

import xlrd
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
    

workbook = xlrd.open_workbook('test.xls')
sheet_names = workbook.sheet_names()
sheet = workbook.sheet_by_index(0)
num_rows = sheet.nrows
num_cols = sheet.ncols
xls_list=[]
for i in range(num_rows):
    row_list=[]
    row_list.append(sheet.cell_value(i, 0))
    row_list.append(sheet.cell_value(i, 1))
    xls_list.append(row_list)
print(xls_list[2][1])
print(len(xls_list))
print(len(xls_list[0]))
print(num_rows)
for k1 in range(183):
    for k2 in range(k1+1,183):
        if xls_list[k1][0]!="" and xls_list[k2][0]!="":
           str1=xls_list[k1][0]
           str2=xls_list[k2][0]
           if similarity(str1, str2)>0.7:
              print(k1)
              xls_list[k1][1]=xls_list[k1][1]+xls_list[k2][1]
              xls_list[k2][0]=""
              xls_list[k2][1]=""

workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('Sheet1')
for m in range(len(xls_list)):
    worksheet.write(m, 0, xls_list[m][0])
    worksheet.write(m, 1, xls_list[m][1])

workbook.save('train.xls')
              
              

# -*- coding: utf-8 -*-
import win32com.client as win32
import sys

para = sys.argv[1]
excel = win32.gencache.EnsureDispatch("Excel.Application")
wb = excel.Workbooks.Open(str(para))
#excel.Visible = True
ws = wb.Worksheets("工作表1")

f = open(r'input.txt')
s = open('output.txt','a+')
s.readline()
text = list()
for line in f:
    str_list = line.split(',')
  
    if str_list[0] == "WRITE":
        ws.Cells(str_list[1],str_list[2]).Value = str_list[3]
        s.write("OK")
        wb.Save()
        
    elif str_list[0] == "Get_Data":
        get_data_list = ws.Range(ws.Cells(1,1), ws.Cells(200,100)).Value
        for i in range (200):
            for j in range(100):
                s.write(str(get_data_list[i][j]) + ",")
            print(";")
        
    else:
        print("para must be 'write' or 'get_data'")

        
#excel 欄位取值
        #ws.Range(A1.Value)
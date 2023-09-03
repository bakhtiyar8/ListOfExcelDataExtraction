from datetime import datetime
import os, sys
import os.path
import time
import shutil
import re
import openpyxl
from openpyxl import Workbook

def siteIdToDic(fileName, write1, write2, write3, write4, write5, write6, write7, write8, write9, write10):
    with open(fileName, "a") as f:
        wr = "{0} {1} {2} {3} {4} {5} {6} {7} {8} {9}".format(write1, write2, write3, write4, write5, write6, write7, write8, write9, write10)
        print(wr)
        f.write("{0}\n".format(wr))


def lpp_th_folder(path_1):
    for root, dirs, files in os.walk(path_1):
        for file in files:
            if file.endswith('.xlsx'):
                print("Try: {0}".format(file))
                path_to_exc = os.path.join(root, file)
                try:
                    book = openpyxl.load_workbook(path_to_exc)
                    sheet = book['Site Audit']
                    a1 = sheet['D7'].value
                    print(a1)
                    sheet2 = book['LTE700 Checklist']
                    a2 = sheet2['E19'].value
                    a3 = sheet2['H19'].value
                    a4 = sheet2['K19'].value

                    sheet3 = book['LTE1800 Checklist_1']
                    a5 = sheet3['E19'].value
                    a6 = sheet3['H19'].value
                    a7 = sheet3['K19'].value

                    sheet4 = book['LTE1800 Checklist_2']
                    a8 = sheet4['E19'].value
                    a9 = sheet4['H19'].value
                    a10 = sheet4['K19'].value

                    siteIdToDic("C://Users//SSV_data2.txt", 
                        'Site id:'+str(a1)+";", 'LTE700 Checklist, 1 sector:'+str(a2)+";", 'LTE700 Checklist, 2 sector:'+str(a3)+";", 
                        'LTE700 Checklist, 3 sector:'+str(a4)+";", 'LTE1800 Checklist, 1 sector:'+str(a5)+";", 'LTE1800 Checklist, 2 sector:'+str(a6)+";", 
                        'LTE1800 Checklist, 3 sector:'+str(a7)+";", 'LTE1800 Checklist, 1 sector:'+str(a8)+";", 'LTE1800 Checklist, 2 sector:'+str(a9)+";", 
                        'LTE1800 Checklist, 3 sector:'+str(a10)+";")
                except Exception as e:
                    print("An exception occurred:", str(e))
                    continue
#Path to folder with excel files
#path_1 = 'C://Users//'
lpp_th_folder(path_2)

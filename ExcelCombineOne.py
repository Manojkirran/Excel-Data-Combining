# combine Multiple Sheets into One Sheet for one file (and we can also Marge selected sheet)

import pandas as pd
import os.path
from os import path
import platform
import os
data = {}
data2 = []

def Mergeallexcel(inputxl,outputxl,*sheetnames,sheet1 ="sheet"):
    filenae = os.path.split(inputxl)
    tmp = filenae[-1].replace(".xlsx", "")
    xls = pd.ExcelFile(inputxl)
    name1 = xls.sheet_names
    for sheet in name1:
        df = pd.read_excel(inputxl, sheet_name=sheet)
        if sheet in data2:
            df = pd.read_excel(inputxl, sheet_name=sheet)
            sheet = sheet + "_" + tmp
            data[sheet] = df
        data2.append(str(sheet))
        data[sheet] = df
#Output file Function

    ostype = platform.system()
    fu = outputxl
    filenae1 = os.path.split(fu)
    if not filenae1[0] == "":
        if ostype == "Windows":
            ful = fu.split("\\")
        else:
            ful = fu.split("/")
    foldercreat = ""
    if not filenae1[0] == "":
        for i in range(0, len(ful) - 1):
             if ostype == "Windows":
                  foldercreat = foldercreat + "/" + ful[i]
        path1 = os.getcwd()
        validpath = path.exists(path1+foldercreat)
        if validpath == False:
            os.makedirs(path1+foldercreat)


    writer = pd.ExcelWriter(outputxl)
    gt = 0
    if len(sheetnames) >= 1:
        for sheet2 in sheetnames:
            if not sheet2 in data:
                print(sheet2,"This sheet name is not reflecting in given file")
            else:
                 df = data[sheet2]
                 tr3 = data.get(sheet2)
                 lenthofdf3 = len(tr3)
                 df.to_excel(writer,sheet_name= sheet1, index =False,startrow=gt)
                 gt = gt + lenthofdf3 + 2
        writer.save()
    else:
         for sheet in data:
             df = data[sheet]
             tr3 = data.get(sheet)
             lenthofdf3 = len(tr3)
             df.to_excel(writer, sheet_name=sheet1, index=False, startrow=gt)
             gt = gt + lenthofdf3 + 2
         writer.save()
    datadu =pd.read_excel(outputxl)
    datadu2=datadu.drop_duplicates()
    writer = pd.ExcelWriter(outputxl)
    datadu2.to_excel(writer, sheet_name=sheet1, index=False)
    writer.save()
outxl =r'Output Path'
Mergeallexcel("File Name","InputPath","Sheetname1","Sheetname2",sheet1 ="What ever Name you want")


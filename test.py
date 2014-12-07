# -*- coding: utf-8 -*- 

"""
注意事项：
1.py文件和xls同一目录  在py中控制不读取最后一个文件（把py文件放在最后）
2.不要事先创建new.xls文件，这是pyexcelerator模块的缺点
3.注意异常控制 并写入日志
"""

import xlrd,pyExcelerator,os,time

def openOneExcel(excelName=''):
    try:
        book = xlrd.open_workbook(excelName)
        return book
    except Exception,e:
        print str(e)


def closeOneExcel():
    #release_resources()
    return

def dir2list(dirName):
    filelists = os.listdir(dirName)
    return filelists


def readOneExcel(excelName=''):
    book = openOneExcel(excelName)
    sheet = book.sheets()[1];
    cellList = sheet.row_values(0)
    return cellList


def writeToOneExcel(dirName='', toExcelName=''):
    filelists = dir2list(dirName)
    i = 0
    w = pyExcelerator.Workbook()
    ws = w.add_sheet('Sheet1')
    for afile in filelists[0:-1]:
        lists = readOneExcel(afile)
        print "%s" % time.ctime()
        print ("read file %s ..." % afile)
        for x in range(len(lists)):
            ws.write(i,x,lists[x])
        i = i+1
        print ("write file %s to new.xls...\n" % afile)
        w.save(toExcelName)
        time.sleep(2)
    print ("end ...")


def main():
    dirName = "./"
    toExcelName = "../new.xls"
    writeToOneExcel(dirName, toExcelName)
    
if __name__ == "__main__":
    main()

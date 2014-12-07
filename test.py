# -*- coding: utf-8 -*- 

"""
注意事项：
1.py文件和xls同一目录  在py中控制不读取最后一个文件（把py文件放在最后）
2.不要事先创建new.xls文件，这是pyexcelerator模块的缺点
3.注意异常控制 并写入日志
"""

#导入需要使用的模块
#xlrd读excel文件，pyExcelerator写excel文件，os遍历目录
import xlrd,pyExcelerator,os,time

'''
功能：打开一个excel文件
参数：等待打开的excel文件名
返回值：打开的excel文件资源
'''
def openOneExcel(excelName=''):
    try:
        book = xlrd.open_workbook(excelName)
        return book
    except Exception,e:
        print str(e)

'''
功能：将制定目录下的文件名转换成列表（os.listdir功能更全面，但是本项目只遍历文件，不需要遍历目录）
参数：目录名
返回值：当前目录下文件组成的列表
'''
def dir2list(dirName):
    filelists = os.listdir(dirName)
    return filelists

'''
功能：读一个excel文件
参数：文件名
返回值：该文件第二个工作表的第一行的数据组成的列表
'''
def readOneExcel(excelName=''):
    book = openOneExcel(excelName)
    sheet = book.sheets()[1];
    cellList = sheet.row_values(0)
    return cellList

'''
功能：读文件，写入excel
参数：目录名，写入的文件名
返回值：无
'''
def writeToOneExcel(dirName='', toExcelName=''):
    filelists = dir2list(dirName)
    # i 局部变量，控制写入的行
    i = 0
    #创建一个excel工作簿
    w = pyExcelerator.Workbook()
    #为工作簿添加sheet1工作表
    ws = w.add_sheet('Sheet1')
    #控制不读取目录中最后一个文件（具体描述见注意事项）
    for afile in filelists[0:-1]:
        lists = readOneExcel(afile)
        print "%s" % time.ctime()
        print ("read file %s ..." % afile)
        #根据单元格坐标写入
        for x in range(len(lists)):
            ws.write(i,x,lists[x])
        i = i+1
        print ("write file %s to new.xls...\n" % afile)
        #每写入一行保存一次
        w.save(toExcelName)
        time.sleep(2)
    print ("end ...")


def main():
    dirName = "./"
    toExcelName = "../new.xls"
    writeToOneExcel(dirName, toExcelName)
    
if __name__ == "__main__":
    main()

#!python3
__author__ = 'panying'

#
# worktime v0.1
#
#  neo4cn@hotmail.com 2015/5/1
#
#  工时管理主程序
#
#  从指定格式的excel文件读取工时记录，并将所有记录写入数据库
#
#  此版本使用sqlite3数据库
#
#  主要处理流程：
#  1、利用内置glob库读取当前目录下excel文件清单
#  2、利用第三方xlrd读取每个excel文件工时记录信息
#  3、利用内置sqlite3库将读取的工时记录写入数据库
#
# neo4cn@hotmail.com 2015/5/5
#  增加评分处理逻辑
#  1、修改Excel模板，增加评分单元格，修改worktime数据表增加评分字段
#  2、根据是否有评分，确定是否导入相应项目的工时数据
# 增加导入时间记录
#

import os
import glob
import datetime
import sqlite3

import xlrd





# 根据员工姓名从person表中取相应工号
def getpersonid(name):
    sql = "select id from person where name ='%s'" % name
    cursor.execute(sql)
    return cursor.fetchone()[0]


# 根据cell所在列从Excel中取工时数据对应的工时日期
def getdate(col):
    col = col - col % 2  # 求工时日期所在单元格，正常与加班工时对应一个工时日期
    date = sheet.cell(3, col).value  # 在excel模板中，工时日期为于第4行
    date = xlrd.xldate_as_tuple(date, 0)  # excel中时间保存为一个浮点数，需要通过xrld中的函数将其转换为YYYY,MM,DD的元组
    date = datetime.datetime(date[0], date[1], date[2]).strftime("%Y%m%d")  #将日期元组转换为YYYYMMDD字符串
    return date


# 根据cell所在列计算是否是加班工时标志，根据当前的模板，单数列为正常工时，双数列为加班工时
def getoverflag(col):
    return col % 2


# 根据cell所在行读取工时数据对应的项目名称，在当前模板中项目名称位于第3列
def getprojectname(row):
    return sheet.cell(row, 2).value


# 根据项目名称从project表中获得项目编号
def getprojectid(name):
    sql = "select id from project where name ='%s'" % name
    cursor.execute(sql)
    return cursor.fetchone()[0]


# 根据cell所在行读取工时数据对应的项目阶段名称，在当前模板中项目阶段名称位于第4列
def getphasename(row):
    return sheet.cell(row, 3).value


# 根据阶段名称从phase表中获得阶段编号
def getphaseid(name):
    sql = "select id from phase where name ='%s'" % name
    cursor.execute(sql)
    return cursor.fetchone()[0]

# 根据cell所在行取评分
def getperformance(row):
    return sheet.cell(row, 19).value


#提示当前工作目录
dir = os.getcwd()
print("工作目录:", dir)

#打开sqlite3数据库
connect = sqlite3.connect("worktime.db")
cursor = connect.cursor()

#获取工作目录中所有xls文件清单
xlsfiles = glob.glob("*.xlsx")

#根据excel模板的格式处理每个excel文件
for file in xlsfiles:

    print("\n工时登记文件:", file)

    #打开excel文件与相应sheet
    workbook = xlrd.open_workbook(file)
    sheet = workbook.sheet_by_name('工时登记')

    #在excel制定单元格中获取工时登记人
    personname = sheet.cell(1, 2).value
    print("工时登记人:", personname)

    personid = getpersonid(personname)

    #根据单元格位置读取工时数据
    for row in range(5, 9 + 1):
        performance = getperformance(row)
        if performance != "" and performance != 0:
            for col in range(4, 17 + 1):
                hour = sheet.cell(row, col).value
                if hour != "" and hour != 0:
                    date = getdate(col)
                    projectid = getprojectid(getprojectname(row))
                    phaseid = getphaseid(getphasename(row))
                    overflag = getoverflag(col)
                    hour = int(hour)
                    print(date, personid, projectid, phaseid, hour, overflag, performance, end=" ")
                    sql = "insert into worktime values ('%s','%s','%s','%s',%s,%s,%s,'%s','%s')" % (
                        date, personid, projectid, phaseid, hour, overflag, performance, "",
                        datetime.datetime.now().strftime("%Y%m%d%H%M%S"))
                    print(sql)
                    try:
                        cursor.execute(sql)
                        connect.commit()
                        print("导入成功.")
                    except sqlite3.IntegrityError:
                        print("与已有记录重复，导入失败.")
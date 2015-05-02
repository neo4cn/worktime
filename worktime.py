__author__ = 'panying'

import os
import glob
import datetime
import sqlite3

import xlrd


def getpersonid(name):
    sql = "select id from person where name ='%s'" % name
    cursor.execute(sql)
    return cursor.fetchone()[0]


def getdate(col):
    col = col - col % 2
    date = sheet.cell(3, col).value
    date = xlrd.xldate_as_tuple(date, 0)
    date = datetime.datetime(date[0], date[1], date[2]).strftime("%Y%m%d")
    return date


def getoverflag(col):
    return col % 2


def getprojectname(row):
    return sheet.cell(row, 2).value


def getprojectid(name):
    sql = "select id from project where name ='%s'" % name
    cursor.execute(sql)
    return cursor.fetchone()[0]


def getphasename(row):
    return sheet.cell(row, 3).value


def getphaseid(name):
    sql = "select id from phase where name ='%s'" % name
    cursor.execute(sql)
    return cursor.fetchone()[0]


dir = os.getcwd()
print("工作目录:", dir)

connect = sqlite3.connect("worktime.db")
cursor = connect.cursor()

xlsfiles = glob.glob("*.xlsx")

for file in xlsfiles:
    print("\n工时登记文件:", file)
    workbook = xlrd.open_workbook(file)
    sheet = workbook.sheet_by_name('工时登记')

    personname = sheet.cell(1, 2).value
    print("工时登记人:", personname)
    personid = getpersonid(personname)

    for row in range(5, 9 + 1):
        for col in range(4, 17 + 1):
            hour = sheet.cell(row, col).value
            if hour != "" and hour != 0:
                date = getdate(col)
                projectid = getprojectid(getprojectname(row))
                phaseid = getphaseid(getphasename(row))
                overflag = getoverflag(col)
                hour = int(hour)
                print(date, personid, projectid, phaseid, hour, overflag, end=" ")
                sql = "insert into worktime values ('%s','%s','%s','%s',%s,%s,'%s')" % (
                date, personid, projectid, phaseid, hour, overflag, "")
                try:
                    cursor.execute(sql)
                    connect.commit()
                    print("保存成功.")
                except sqlite3.IntegrityError:
                    print("与已有记录重复，保存失败.")
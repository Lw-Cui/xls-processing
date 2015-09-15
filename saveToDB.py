#!/usr/bin/python
import MySQLdb
import string
import xlrd

def extract():
	sheet = xlrd.open_workbook("grade.xls").sheet_by_index(0)
	for rx in range(1, sheet.nrows):
		row = sheet.row(rx)
		info = {}
		info["name"] = row[3].value
		info["course"] = row[7].value
		info["final"] = row[10].value
		#info["resit"] = row[11].value
		info["resit"] = -1
		yield info

connect = MySQLdb.connect(host="localhost",
	port=3306, user="root", passwd="greatclw", db="grade", charset="utf8", use_unicode=True)
cur = connect.cursor()
#cur.execute(
	#"CREATE TABLE info(name varchar(20), course varchar(30), final integer, resit integer) character set utf8")
for student in extract():
	#cur.execute(u"INSERT INTO info VALUES(\'%s\', \'%s\', %s, %s)" % (student["name"], student["course"], student["final"], student["resit"]))
	print u"INSERT INTO info VALUES(\'%s\', \'%s\', %s, %s)" % (student["name"], student["course"], student["final"], student["resit"])
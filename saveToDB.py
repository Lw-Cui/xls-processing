#!/usr/bin/python
# -*- coding: utf-8 -*-
from __future__ import division
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
		info["resit"] = row[11].value

		if info["final"] == u'':
			info["final"] = 0.0

		if info["final"] >= 60.0 and info["resit"] == u'':
			info["resit"] = 100.0
		elif float(info["final"]) < 60.0 and info["resit"] == u'':
			info["resit"] = 0.0

		yield info

connect = MySQLdb.connect(host="localhost",
	port=3306, user="root", passwd="greatclw", db="grade", charset="utf8", use_unicode=True)
cur = connect.cursor()
cur.execute("DROP TABLE IF EXISTS info")
cur.execute(
	"CREATE TABLE info(name VARCHAR(20), course VARCHAR(30), final DOUBLE, resit DOUBLE) character set utf8")
for student in extract():
	sql = "INSERT INTO info VALUES('%s', '%s', %s, %s)" % (student["name"], student["course"], student["final"], student["resit"])
	print sql
	cur.execute(sql)
connect.commit()

course = [u'大学物理Ⅰ', u"线性代数与空间解析几何I", u"通用英语", u'微积分I', u'微积分II', u"离散数学"]
for query in course:
	sum = "select * from info where course = '%s'" % query
	final = "select * from info where course = '%s' and final < 60" % query
	resit = "select * from info where course = '%s' and final < 60 and resit < 60" % query

	print "__________ %s __________" % query
	print "正考不及格率:  %.2f%%" % (cur.execute(final) / cur.execute(sum) * 100)
	print "补考后不及格率:  %.2f%%" % (cur.execute(resit) / cur.execute(sum) * 100)
	print "\n"

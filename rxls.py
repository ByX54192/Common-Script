# -*- coding: UTF-8 -*-
import xlrd
from datetime import date,datetime
#python读取excel表中单元格的内容返回的有5种类型，即ctype:
# ctype: 0 empty ,1 string ,2 number,3 date,4 boolean,5 error
#读取的文件名
rfile='test1.xlsx'

def read_excel():
	wb = xlrd.open_workbook(filename=rfile)
	sheet_list=wb.sheet_names()
	sheet1=wb.sheet_by_index(0) #通过索引获取表格
	sheet2=wb.sheet_by_name('工资') # 通过名字获取表格
	#print(sheet1,sheet2) 此处打印sheet两个变量的内存地址
	#print(sheet1.name,sheet1.nrows,sheet1.ncols)
	#rows=sheet1.row_values(1) #获取行内容
	#cols=sheet1.col_values(0) #获取列内容
	#print(rows,cols)

	#print(sheet2.name,sheet2.nrows,sheet2.ncols)

	# 获取表格里的内容，三种方式
	# print(sheet1.cell(1,2).value) #即打印第1行第0列的内容
	# print(sheet1.cell_value(1,2))
	# print(sheet1.row(1)[2].value)

	#print(sheet1.cell(1,2).ctype)# 即 ctype的值
	# 处理时间格式用xlrd的模块处理
	date_value = xlrd.xldate_as_tuple(sheet1.cell_value(1,2),wb.datemode)
	print(date(*date_value[:3])) #第一种时间格式
	print(date(*date_value[:3]).strftime('%Y/%m/%d'))
  
if __name__=="__main__":
	read_excel()
# -*- coding: UTF-8 -*-
# author:by Seven
import xlwt
#设置表格样式
def set_style(name,height,bold=False):
	style = xlwt.XFStyle()
	font = xlwt.Font()
	font.name = name
	font.bold = bold
	font.color_index = 4
	font.height = height
	style.font =font
	return style
#写Excel 
def write_excel():
	f = xlwt.Workbook()
	sheet1 = f.add_sheet("网站统计",cell_overwrite_ok=True)
	row0 = ["网站","类型","访问量","Rank"]
	colum0 = ["百度","新浪","网易","天涯论坛","知乎"]
	#write_merge()用法；sheet1.write(1,2,3,3,'demo') 具体查询官方文档
	  
	#写第一行
	for i in range(0,len(row0)):
		sheet1.write(0,i,row0[i],set_style('Time New Roman',220,True))
	#写第一列
	for j in range(0,len(colum0)):
		sheet1.write(j+1,0,colum0[j],set_style('Time New Roman',220,True))
	# sheet1.write()
		
	f.save("test.xls")

if __name__=="__main__":
 	write_excel()

__author__ = 'furong'
#-*- coding:utf-8 -*-
#------xlsxwriter模块向xlsx文件中写入数据，xlrd模块读取xlsx文件内容
#导入xlsxWriter模块
import xlsxwriter as xlw
#----------基本使用
#创建xlsx文件
workbook=xlw.Workbook('test.xlsx')
#创建sheet
worksheet=workbook.add_worksheet('sheetname1')
#向a1写入数据
#worksheet.write('A1','hi world')

#----------excel计算公式
#待添加数据
data=(
    ['math',100],
    ['english',90],
    ['physic',60]
)
#从0行0列开始添加数据
row=0
col=0
for item,value in (data):
    worksheet.write(row,col,item)
    worksheet.write(row+1,col,value)
    col+=1
worksheet.write(0,col,'totel')
worksheet.write(1,col,'=SUM(A2:C2)')


#---------------创建excel图表
worksheet2=workbook.add_worksheet('sheetname2')
#创建新的chart对象
chart=workbook.add_chart({'type':'column'})
#设置chart的标题，x轴标题，y轴标题
chart.set_title({'name':'统计数据'})
chart.set_x_axis({'name':'学号'})
chart.set_y_axis({'name':'分数'})
#设置格式，粗体
bold=workbook.add_format({'bold':1})
#待添加数据
data2 = [
    [1, 2, 3, 4, 5],
    [2, 4, 6, 8, 10],
    [3, 6, 9, 12, 15],
]
#表格标题数据
sheet2title=['sam','tom','lili']
#写入标题，设置为粗体
worksheet2.write_row('A1',sheet2title,bold)
#A列依次填充【1,2,3,4,5】，B列依次填充【2,4,，6,8,10】，C列依次填充【3,6,9,12,15】
worksheet2.write_column('A2',data2[0])
worksheet2.write_column('B2',data2[1])
worksheet2.write_column('C2',data2[2])

#给chart添加数据
chart.add_series({'values': '=sheetname2!$A$1:$A$5'})
chart.add_series({'values': '=sheetname2!$B$1:$B$5'})
chart.add_series({'values': '=sheetname2!$C$1:$C$5'})

#将chart对象添加到sheet中
worksheet2.insert_chart('A7',chart)
workbook.close()

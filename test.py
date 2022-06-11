'''
Author: yangrongxin
Date: 2022-06-11 11:26:11
LastEditors: yangrongxin
LastEditTime: 2022-06-11 11:26:14
'''
import readAndWrite
openexcel = readAndWrite.ReadExcel(file_name='./demo1.xls',sheet_id=0)
openexcel2 = readAndWrite.ReadExcel(file_name='./demo2.xls',sheet_id=0)
write_excel = readAndWrite.WriteExcel('第一张表')

for i in range(1,openexcel.get_lines()):
    #将目标表格的姓名拷贝在整理的表格中
    write_excel.write_values(i, 0, openexcel.get_value(i,0))
    #将身份证号码中的生日提取出来
    birthday = openexcel.get_value(i,1)
    #将生日写入到目标表格中
    write_excel.write_values(i,1,birthday)

write_excel.add_sheet('第二张表')

for i in range(1,openexcel2.get_lines()):
    #将目标表格的姓名拷贝在整理的表格中
    write_excel.write_values(i, 0, openexcel2.get_value(i,0))
    #将身份证号码中的生日提取出来
    birthday = openexcel2.get_value(i,1)
    #将生日写入到目标表格中
    write_excel.write_values(i,1,birthday)

write_excel.save_file(filename="total.xls")
'''
Author: yangrongxin
Date: 2022-06-11 10:35:18
LastEditors: yangrongxin
LastEditTime: 2022-06-11 11:15:58
'''
import xlwt
import xlrd

class ReadExcel:
    # 初始化函数 确定表名称 和 sheet index
    def __init__(self, file_name=None, sheet_id=None):
        if file_name:
            self.file_name = file_name
            self.sheet_id = sheet_id
        else:
            self.file_name = 'default.xls'
            self.sheet_id = 0
        self.data = self.get_data()
    # 获取表格数据
    def get_data(self):
        data = xlrd.open_workbook(self.file_name)
        tables = data.sheets()[self.sheet_id]
        return tables
    # 获取表格中具体行列数据
    def get_value(self, row, col):
        return self.data.cell_value(row, col)
    # 获取表格总列数
    def get_lines(self):
        return self.data.nrows
    # 获取表格总行数
    def get_cols(self):
        return self.data.ncols


class WriteExcel:
    # 初始化需要生成的表名称
    def __init__(self, sheet_name = None):
        if sheet_name:
            self.sheet_name = sheet_name
        else:
            self.sheet_name = 'sheet1'
        self.instance = xlwt.Workbook()
        self.worksheet = self.instance.add_sheet(sheetname = self.sheet_name)

    # 增加一个新的sheet添加数据
    def add_sheet(self, sheet_name = None):
        self.worksheet = self.instance.add_sheet(sheetname = sheet_name)

    # 向目标行写入数据
    def write_values(self, row, col, values):
        self.worksheet.write(row, col, values)

    # 存储表格
    def save_file(self, filename = None):
        if filename:
            self.filename = filename
        else:
            self.filename = 'default.xls'
        self.instance.save(self.filename)
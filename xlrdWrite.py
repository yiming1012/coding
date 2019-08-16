# -*- coding: utf-8 -*-
import xlrd
from xlutils.copy import copy


filename = 'a.xls'# 文件名
rb = xlrd.open_workbook(filename, formatting_info=True)
# formatting_info=True: 保留原数据格式
wb = copy(rb) 			# 复制页面
ws = wb.get_sheet(0) 	# 取第一个sheet
# ----- 按(row, col, str)写入需要写的内容 -------
for i in range(96):
	ws.write(1, 5 + i, 1)
# ----- 按(row, col, str)写入需要写的内容 -------
wb.save(filename)

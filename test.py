# -*- coding: utf-8 -*-
# @Author: fullback555
# @Date:   2020-05-23 16:17:11
# @Last Modified by:   fullback555
# @Last Modified time: 2020-05-29 23:56:52

# from config import MyPath
# import os

# path_wk = MyPath.get_cur_path('wkhtmltopdf.app')
# if not os.path.exists(path_wk):
# 	# MyPath.copy_file(filename='wkhtmltopdf')
# 	print('不存在：{}'.format(path_wk))

# 判断数组是否连续
def isContinuousArray(alist):
	min = 9999
	max = 0
	for i in range(len(alist)):
		if alist[i] == 0:
			continue
		elif alist[i] < min:
			min = alist[i]
		elif alist[i] > max:
			max = alist[i]
	if max - min < len(alist)-1:
		return True
	else:
		return False

a = [0,1,2,3,4]
b = [3,4,6]
print(isContinuousArray(a))
print(isContinuousArray(b))

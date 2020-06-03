# -*- coding: utf-8 -*-
# @Author: wenlong
# @Date:   2020-05-07 23:03:03
# @Last Modified by:   fullback555
# @Last Modified time: 2020-05-30 20:16:40

# 图片处理
from PIL import Image
import numpy as np
import os


class MyImage():
	"""处理百度文库下载的图片"""
	def __init__(self, picpath):
		self.picpath = picpath
		try:
			self.image = Image.open(picpath)
		except Exception as e:
			self.image = None
			print('打开图片失败，原因：{}'.format(e))
		

	# 判断图片是否有空白
	def has_blank(self, image, picmode='P'):
		image = self.convert_mode(image, picmode)
		img = np.array(image)
		print(img[90:100])
		h, w = img.shape[-5:]  # 图片的宽和高
		pic_blank = [i for i in range(h) if sum(img[i]) >= 254 * w and np.var(img[i]) <= 1]
		if len(pic_blank) >= 5:
			print('it has blank')
			return True
		return False

	# 转换图片模式并保存, 接收参数为图片路径
	def convert_save_mode(self, picpath, mode='L'):
		# PIL中有九种不同模式。分别为1，L，P，RGB，RGBA，CMYK，YCbCr，I，F。
		image = Image.open(picpath)
		image_new = image.convert(mode)
		file, suffix = os.path.splitext(self.picpath)  # 将文件名和扩展名分开
		filename = file + '_' + mode + suffix  # 拼接新路径
		image_new.save(filename)
		print('{}转化为{}成功。'.format(image.mode, image_new.mode))
			
	# 转换图片模式, 接收参数为图片路径或image对象
	def convert_mode(self, image, mode='P'):
		# PIL中有九种不同模式。分别为1，L，P，RGB，RGBA，CMYK，YCbCr，I，F。
		if isinstance(image, str):
			image = Image.open(image)
		# print(image.mode, image.format, image.size)
		if image.mode != mode:  # 如果图像模式不是P模式,那么都转化为P模式
			image = image.convert(mode)
			# print(image.mode, image.format, image.size)
		return image

	# 判断是否为全白图片
	def is_blank(self, image, box=None):
		if not box:
			img = image
		else:
			img = image.crop(box)
		extrema = img.convert("L").getextrema()
		print(extrema)
		if extrema == (0, 0):  # all black
			pass
		elif extrema == (255, 255):  # all white
			return True
		return False

	# 判断数组是否连续，至少5个元素
	def isContinuousArray(self, alist):
		min = 9999
		max = 0
		for i in range(len(alist)):
			if alist[i] == 0:
				continue
			elif alist[i] < min:
				min = alist[i]
			elif alist[i] > max:
				max = alist[i]
		if max - min <= len(alist)-1 and len(alist) >= 5:
			return True
		else:
			return False

	# 获取图片的真实范围(主要是高度上), 模式mode分为垂直方向v，水平方向h，实际上就是按垂直或水平方向切割
	def get_picrange(self, image, cutmode='v', picmode='RGBA'):
		image = self.convert_mode(image, picmode)

		w, h =image.size
		print('h:{}, w:{}'.format(h,w))

		# 经过分析，百度文库doc文档同一页上面所有图片会合成放在同一张图上，中间有高度5个像素高度（透明色）隔开。
		if cutmode == 'v':
			pic_blank = [i for i in range(h) if image.getpixel((0, i))[3] == 0 and image.getpixel((int(w/2), i))[3] == 0]
			pic_true = set([i for i in range(h)]) - set(pic_blank)  # 集合去重，求差集
		elif cutmode == 'h':
			pic_blank = [i for i in range(w) if image.getpixel((i, 0))[3] == 0 and image.getpixel((i, int(h/2)))[3] == 0]
			pic_true = set([i for i in range(w)]) - set(pic_blank)  # 集合去重，求差集
		pic = sorted(list(pic_true))
		print(pic)

		if not pic:
			print('pic', pic)
			return []

		picrange = []
		start = pic[0]
		start_i = 0
		for i in range(len(pic)-1):
			# if pic[i+1] - pic[i] >= 5 and pic[i] - start > 4:
			if pic[i+1] - pic[i] >= 5:
				if self.isContinuousArray(pic[start_i: i+1]):
			# if pic[i+1] - pic[i] >= 5 and self.isContinuousArray(pic[start_i: i+1]):  # 用5个像素隔开为有效，分辨率为5个像素
					print('start_i:{}, i:{}'.format(start_i, i))
					print(pic[start_i: i+1])
					picrange.append((start, pic[i]+1))
					start = pic[i+1]
					start_i = i + 1
				else:
					start = pic[i+1]
					start_i = i + 1
		if pic[-1] - start >= 5 and self.isContinuousArray(pic[start_i: ]): # 最后面的真实图片
			picrange.append((start, pic[-1]+1))
		print('真实图片所有区域{}:{}'.format(cutmode, picrange))

		return picrange

	# 获取图片的真实范围(主要是高度上), 模式mode分为垂直方向v，水平方向h，实际上就是按垂直或水平方向切割
	# def get_picrange(self, image, cutmode='v', picmode='P'):
	# 	image = self.convert_mode(image, picmode)
	# 	img = np.array(image)
	# 	# print('np图片数组信息：', img.shape, img.ndim, img.dtype)
	# 	if cutmode != 'v':  # 非垂直模式，即水平模式
	# 		img = img.T  # 矩阵转置，变成垂直方向
	# 	h, w = img.shape[:2]  # 图片的宽和高
	# 	print(img[-5:])
	# 	# 经过分析，百度文库doc文档同一页上面所有图片会合成放在同一张图上，中间有高度5个像素高度（透明色）隔开。
	# 	# 有时图片上部会是一大片空白区域。隔开的这5个像素值通常为254或255（图像P模式）.
	# 	# pic = [i for i in range(h) if sum(img[i]) < 254 * w]  # np.var(img[i]) > 1
	# 	# pic = [i for i in range(h) for j in img[i] if j < 254]

	# 	pic_blank = [i for i in range(h) if sum(img[i]) >= 254 * w or sum(img[i]) == 128 * w and np.var(img[i]) <= 1]
	# 	pic_true = set([i for i in range(h)]) - set(pic_blank)  # 集合去重，求差集
	# 	pic = sorted(list(pic_true))

	# 	if not pic:
	# 		print('pic', pic)
	# 		return []

	# 	picrange = []
	# 	start = pic[0]
	# 	for i in range(len(pic)-1):
	# 		if pic[i+1] - pic[i] > 5 and pic[i] - start >= 4:  # 用5个像素隔开为有效，分辨率为5个像素
	# 			picrange.append((start, pic[i]+1))
	# 			start = pic[i+1]
	# 	if pic[-1] - start > 5: # 最后面的真实图片
	# 		picrange.append((start, pic[-1]+1))
	# 	print('真实图片所有区域{}:{}'.format(cutmode, picrange))

	# 	return picrange

	# 获取图片的真实区域
	def get_boxs(self, image, picrange, cutmode='v'):
		boxs = []  # 图像区域
		for y1, y2 in picrange:
			# box坐标为（左，上，右，下）。 PIL使用的坐标系的左上角为（0，0）
			if cutmode == 'v':  # 垂直模式
				box = (0, y1, image.size[0], y2)
			else:  # 水平模式
				box = (y1, 0, y2, image.size[1])
			# 如果不是全白图片
			if not self.is_blank(image, box):
				boxs.append(box)
		# print(boxs)
		return boxs

	# 根据空白区域分割图片
	def cut_image(self):
		print(self.picpath)
		picpaths = []
		# if not self.has_blank(self.image, picmode='P'):
		# 	picpaths = [self.picpath]
		# 	return picpaths
		# 垂直切割
		picrange = self.get_picrange(self.image, cutmode='v', picmode='RGBA')
		boxs = self.get_boxs(self.image, picrange, cutmode='v')

		n = 1
		for i, box in enumerate(boxs):  # 垂直切割
			# 水平切割
			region = self.image.crop(box)  # 从图像复制子矩形
			picrange_h = self.get_picrange(region, cutmode='h', picmode='RGBA')
			boxs_h = self.get_boxs(region, picrange_h, cutmode='h')

			for j, box_h in enumerate(boxs_h):  # 水平切割
				region_h = region.crop(box_h)
				picrange_v = self.get_picrange(region_h, cutmode='v', picmode='RGBA')
				print('picrange_v:', picrange_v)
				boxs_v = self.get_boxs(region_h, picrange_v, cutmode='v')

				for k, box_v in enumerate(boxs_v):  # 垂直切割
					region_v = region_h.crop(box_v)
					file, suffix = os.path.splitext(self.picpath)  # 将文件名和扩展名分开
					filename = file + '_' + str(n) + suffix  # 拼接新路径
					region_v.save(filename)
					n += 1
					picpaths.append(filename)  # 保存图片路径
					print(f'已切割图片保存为：{filename}, {region_h.size}')
		return picpaths


if __name__ == '__main__':
	pass
	
	picpath = '百度文库下载/doc/images/任务4-DT9205A数字万用表装配与调试/41.png'
	pic = MyImage(picpath)
	pic.cut_image()

	# print(pic.is_blank(pic.image))
	# pic.get_picrange(pic.image, cutmode='h')
	# pic.convert_save_mode(pic.picpath, 'P')
	# pic.has_blank(pic.image)

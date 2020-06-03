#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''
@File       :  main.py    
@Contact    :  fullback555@163.com
@Modify Time:  2020/5/17 上午9:26 
@Author     :  pymark0202
@Desc       :  None
'''
from PyQt5.QtWidgets import (QMainWindow, QApplication, QMessageBox, QFileDialog)
from PyQt5.QtCore import QThreadPool
from PyQt5 import QtGui
import sys, os

from mainWin import Ui_MainWindow
from getall import GetAll
from wenkuall import WenKuSpider
from config import *


class MyWindow(QMainWindow, Ui_MainWindow):
	"""主窗口"""

	# 初始化
	def __init__(self):
		super(MyWindow, self).__init__()
		self.setupUi(self)
		self.progressBar.setVisible(False)  # 隐藏进度条

		# 显示首页
		self.homepage()

		self.threadpool = QThreadPool()  # 进程池
		print("Multithreading with maximum %d threads" % self.threadpool.maxThreadCount())

		self.init()

	# 数据初始化
	def init(self):
		# 配置文件初始化
		initConfig()
		self.downloadMode = getConfig('baiduwenku', 'download_mode')  # 下载方式
		# 关于标签
		self.labelAboutLink.setText("<a href='https://github.com/M010K/BaiduWenkuSpider'>M010K的GitHub项目BaiduWenkuSpider</a>")
		self.labelAboutLink.setOpenExternalLinks(True)  # 打开允许访问超链接,默认是不允许
		# 下载完成信号
		self.finished = True
		

	# 一、首页
	def homepage(self):
		# 跳转到page1
		self.stackedWidget.setCurrentIndex(0)

	# 二、设置
	def setting(self):
		# 跳转到page2
		self.stackedWidget.setCurrentIndex(1)
		# 下载方式
		if getConfig('baiduwenku', 'download_mode') == 'mode1':
			self.rBtnMode1.setChecked(True)  # 默认为下载方式一
		else:
			self.rBtnMode2.setChecked(True)

	# 下载方式选择，点击rBtn自动执行
	def selectDownloadMode(self):
		if self.rBtnMode1.isChecked():
			self.downloadMode = 'mode1'
			setConfig('baiduwenku', 'download_mode', 'mode1')
			print('方式一')
		elif self.rBtnMode2.isChecked():
			self.downloadMode = 'mode2'
			setConfig('baiduwenku', 'download_mode', 'mode2')
			print('方式二')

	# 下载按钮
	def download(self):
		# 设置默认下载目录
		savepath = getConfig('path', 'savepath')
		if not savepath:
			filepath = self.setDownloadPath()
			if not filepath:
				return
		savepath = os.path.join(getConfig('path', 'savepath'), '百度文库下载')

		# 如果未完成下载，就等待
		if not self.finished:
			QMessageBox.information(self, '提示', '当前百度文库下载未完成，请稍等！')
			return

		# 获取百度文库url
		url = self.lineEditWenku.text()
		if not self.isBaiduWenKu(url):
			QMessageBox.information(self, '提示', '请输入正确的百度文库链接')
			return None
		self.getDoc(url, savepath)
		self.finished = False  # 下载完成信号

	# 设置默认下载目录
	def setDownloadPath(self):
		filepath = QFileDialog.getExistingDirectory(self, '请选择默认下载目录……', '.')  # 第三个参数'.'代表程序所在目录
		if filepath:
			# 保存到config
			setConfig('path', 'savepath', filepath)
			QMessageBox.information(self, '提示', '已设置默认下载目录为：{}'.format(filepath))
		return filepath

	# 下载文档
	def getDoc(self, url, savepath):
		if self.downloadMode == 'mode1':
			self.doc = GetAll(url, savepath)
		else:
			self.doc = WenKuSpider(url, savepath)
		# doc.signals.result.connect(self.print_result)  # 执行的结果
		self.doc.signals.finished.connect(self.task_finished)  # 执行完毕
		self.doc.signals.condition.connect(self.task_condition)  # 当前状态
		self.doc.signals.progress.connect(self.task_progress)  # 进度
		# Execute，开始线程。
		self.threadpool.start(self.doc)
		# 显示当前状态
		self.labelStatus.setText('正在爬取网页……')
		self.progressBar.setVisible(True)  # 显示进度条
		self.progressBar.setValue(0)  # 设置进度条初始值

		
	# 任务执行完毕，线程结束
	def task_finished(self, s):
		print('complete:', s)
		self.finished = True  # 下载完成信号

		if s == str(False):
			print('下载失败')
			QMessageBox.information(self, '提示', '下载失败，请检查网络重试。')
			return None
		QMessageBox.information(self, '提示', '下载完成。\n文件保存到：{}'.format(s))

	# 当前执行状态
	def task_condition(self, s):
		print(s)
		self.labelStatus.setText(s)

	# 接收进度信号，打印出来
	def task_progress(self, n):
		print("{} done".format(n))
		self.progressBar.setValue(n)

	# 检查是否是百度文库的url
	def isBaiduWenKu(self, url):
		flag = 'https://wenku.baidu.com/view/'
		if flag in url:
			return True
		return False

	# 关于
	def about(self):
		QMessageBox.information(self, '关于', '百度文库下载工具，严禁用于商业用途。如有侵权，请告之删除。\
		作者：pymark0202，联系邮箱：fullback555@163.com。')

	# 重写，显示加载背景图片
	def paintEvent(self, event):
		# 如果资源文件目录中存在bg.jpg，则复制到应用程序目录中来。
		if not os.path.exists(MyPath.get_cur_path('bg.jpg')):
			MyPath.copy_file(filename='bg.jpg')

		painter = QtGui.QPainter(self)
		painter.drawRect(self.rect())

		# 获取当前应用程序目录下的背景图片，并设置为应用程序的背景
		filename = MyPath.get_cur_path('bg.jpg')
		if os.path.exists(filename):  # 如果程序根目录存在背景图片
			pixmap = QtGui.QPixmap(filename)  # 换成自己的图片
			painter.drawPixmap(self.rect(), pixmap)
			return


if __name__ == "__main__":
	app = QApplication(sys.argv)
	myshow = MyWindow()
	myshow.show()
	sys.exit(app.exec_())
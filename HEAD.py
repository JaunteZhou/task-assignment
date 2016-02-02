#-*- coding: utf8 -*-

# Python 2.7.6
# Author ： JaunteZhou（周潮）
# Time ： 2016.02.01
# Type : V1.0.20160201
# File : Class Parameter	『宏定义』变量类文件

import platform
import sys,os

############################
# # # 获得excel文件全称 # # #
############################
def Get_XLS_Full_Name(root, xls_name) :
	file_name = root + '/' + xls_name + r'.xls'
	if ( os.path.exists( file_name ) ) :
		
		#print file_name
		
		return file_name
	elif ( os.path.exists( file_name + 'x' ) ) :
		file_name += 'x'
		
		#print file_name
		
		return file_name
	else :
		print "FILE \'", file_name, "\' or \'", file_name, "x\' is not exists !!!"
		return False
############################
############################
############################

########################
# # # 公共宏定义变量 # # #
########################
class Parameter :
	def __init__(self) :
		self.NUM_ROWS = 4;		#行数，代表『一天4个时间段（大课）』
		self.NUM_COLS = 5;		#列数，代表『一周5天』
		self.NUM_MEMBERS = 12;	#组员数量，12人
		self.NUM_GROUP = 6;		#组数，6组
		# platform.uname()      
		# 包含上面所有的信息汇总，uname_result(system='Windows', node='hongjie-PC',
		# release='7', version='6.1.7601', machine='x86', 
		# processor='x86 Family 16 Model 6 Stepping 3, AuthenticAMD')
		uname_result = platform.uname()
		self.SYSTEM = platform.system()		# 'Linux','Windows','Darwin'

		self.FILE_DIR = self.__file_dir()
		self.DATA_DIR = self.FILE_DIR + '/data'

	#获取脚本文件的当前路径
	def __file_dir(self):
		#获取脚本路径
		path = sys.path[0]

		#判断为脚本文件还是py2exe编译后的文件，
		#如果是脚本文件，则返回的是脚本的目录，
		#如果是py2exe编译后的文件，则返回的是编译后的文件路径
		if os.path.isdir(path):
			return path
		elif os.path.isfile(path):
			return os.path.dirname(path)

########################
########################
########################
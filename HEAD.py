#-*- coding: utf8 -*-

# Python 2.7.6
# Author ： JaunteZhou（周潮）
# Time ： 2016.02.01
# Type : V1.0.20160201
# File : Class Parameter	『宏定义』变量类文件

import platform
import sys,os
import shutil
import datetime

############################
# # # 获得excel文件全称 # # #
############################
def Get_XLS_Full_Name(root, xls_name) :
	file_name = root + '/' + xls_name + r'.xls'
	if ( os.path.exists( file_name ) ) :
		return file_name
	elif ( os.path.exists( file_name + 'x' ) ) :
		file_name += 'x'
		return file_name
	else :
		print "FILE \'", file_name, "\' or \'", file_name, "x\' is not exists !!!"
		return False
############################
############################
############################

############################
# # # 转移temp文件 # # #
############################
def Remove_Temp_File(PARA, file_list) :
	# file_name = root + '/' + xls_name + r'.xls'
	if not os.path.exists( PARA.TEMP_DIR ):
		os.mkdir(PARA.TEMP_DIR)

	for fn in file_list:
		fn_path = fn.split(os.sep)
		fn_first = fn_path[-1].split('.')
		fn_split = fn_first[-2].split('_')

		# # # 如果不在组别序号范围内，则跳过
		if fn_split[-1] == 'temp':
			# # # 获得日期
			date = datetime.date.today()
			s_date = str(date).replace('-','')

			new_fn = PARA.TEMP_DIR + os.sep + fn_first[-2] + '_' + s_date + '.' + fn_first[-1]

			os.rename( fn, new_fn )
	# end for

	
############################
############################
############################

############################
# # # 获得excel文件列表 # # #
############################
def Get_XLS_Fnlist(PARA) :

	# # # '/Users/JaunteZhou/Documents/Github/task-assignment/data'
	fn_list =  os.listdir(PARA.DATA_DIR)

	use4_fn_list = []
	# PARA.NUM_GROUP = 6
	for i in range( PARA.NUM_GROUP ):
		use4_fn_list += ['']

	#print "use4_fn_list:",use4_fn_list
	#print fn_list
	for fn in fn_list:
		fn_first = fn.split('.')
		fn_split = fn_first[0].split('_')

		# # # 如果不在组别序号范围内，则跳过
		if fn_split[0]<'1' or fn_split[0]>str(PARA.NUM_GROUP):
			print "not num !!!"
			continue

		# # # 提取组别序号
		i = int(fn_split[0])
	
		# # # 如果是长期表，则比较时间先后，进行选择
		if fn_split[1] == 'long':
			if use4_fn_list[i-1] == '':
				use4_fn_list[i-1] = fn_split[2]
			elif fn_split[2] > use4_fn_list[i-1]:
				use4_fn_list[i-1] = fn_split[2]
		# # # 如果是临时的，直接选用临时的
		elif fn_split[1] == 'temp':
			# # # '20160320' < 'temp',即使之后有进行比较，也会选择'temp'
			use4_fn_list[i-1] = 'temp'
		# end if
	# end for

	print "use4_fn_list:",use4_fn_list

	for i in range( PARA.NUM_GROUP ):
		file_name = PARA.DATA_DIR + os.sep + str(i+1) + '_' 
		if use4_fn_list[i] != 'temp':
			file_name += 'long_'

		file_name += use4_fn_list[i] + r'.xls'
		print file_name

		if ( os.path.exists( file_name ) ) :
			use4_fn_list[i] = file_name
		elif ( os.path.exists( file_name + 'x' ) ) :
			file_name += 'x'
			use4_fn_list[i] = file_name
		else :
			print "FILE \'", file_name, "\' or \'", file_name, "x\' is not exists !!!"
			use4_fn_list[i] = ''
	# end for
	return use4_fn_list
############################
############################
############################

########################
# # # 公共宏定义变量 # # #
########################
class Parameter :
	def __init__(self, members=11, group=6, rows=4, cols=5 ) :
		self.NUM_ROWS = rows;		#行数，代表『一天4个时间段（大课）』
		self.NUM_COLS = cols;		#列数，代表『一周5天』
		self.NUM_MEMBERS = members;	#组员数量，12人
		self.NUM_GROUP = group;		#组数，6组
		# platform.uname()      
		# 包含上面所有的信息汇总，uname_result(system='Windows', node='hongjie-PC',
		# release='7', version='6.1.7601', machine='x86', 
		# processor='x86 Family 16 Model 6 Stepping 3, AuthenticAMD')
		uname_result = platform.uname()
		self.SYSTEM = platform.system()		# 'Linux','Windows','Darwin'

		self.FILE_DIR = self.__file_dir()
		self.DATA_DIR = self.FILE_DIR + os.sep + 'data'
		self.TEMP_DIR = self.DATA_DIR + os.sep + 'temp'
		self.CCT_DIR = self.DATA_DIR + os.sep + 'CCT'

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
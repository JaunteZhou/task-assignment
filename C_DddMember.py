#-*- coding: utf8 -*-

# Python 2.7.6
# Author ： JaunteZhou（周潮）
# Time ： 2016.02.01
# Type : V1.0.20160201
# File : Class DddMember	督导队成员类文件

import HEAD

PARA = HEAD.Parameter();

class DddMember :
	def __init__(self, name, group, id) :
		self.name = name
		self.group = group
		self.id = id 			# id = 组号(g) * 100 + 组内编号(No)

		self.sum_checkclass = 0			#查课数量
		self.day_checkclass = 0 		#查课日
		self.period_checkclass = 0 		#查课时间段

		self.num_emptyclass = 0		#本周空课数量		empty class num
		self.class_table = []

		for i in range(0, PARA.NUM_COLS) :		#依次生成5天
			self.class_table += [ [] ]
			for j in range(0, PARA.NUM_ROWS) :	#在某1天中加入4个时间段
				self.class_table[i] += [""]
	
	def Insert_ClassTable(self, week, ct) :				
	#ct为传入的空课表，值为『0』『1』『2』『3』
		for i in range(0, PARA.NUM_COLS) :
			for j in range(0, PARA.NUM_ROWS) :
				self.class_table[i][j] = ct.cell( j+1, i+1 ).value	# # # 读取由（1，1）至（5，4）数据
				if (self.class_table[i][j] == week) :
					self.num_emptyclass += 1
				elif (self.class_table[i][j] == 3) :
					self.num_emptyclass += 1

		#self.Show_Info()

	def Show_Info(self) :
		print "Name :", self.name
		print "Group :", self.group
		print "ID :", self.id
		print "Sum of Check Class (this week):", self.sum_checkclass
		print "Num of Empty Class (this week):", self.num_emptyclass
		print "Class Table :", self.class_table
		print ""

# test :
# m1 = DddMember("LYJ", 2)
# m1.Show_Info()
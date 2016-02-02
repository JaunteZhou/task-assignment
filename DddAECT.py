#-*- coding: utf8 -*-

# Python 2.7.6
# Author ： JaunteZhou（周潮）
# Time ： 2016.02.01
# Type : V1.0.20160201
# File : DddAECT (Dudaodui Arrange Empty Class Table)	督导队排空课表

# xlrd 版本 0.9.3
import os
import xlrd, xlwt
import C_DddMember
import C_EmptyClassTable
import HEAD

from datetime import datetime

#################
# # # 主函数 # # #
#################
if __name__ == "__main__" :
	# 常用变量，用类代替宏定义
	PARA = HEAD.Parameter()
		# NUM_ROWS = 4;		#行数，代表『一天4个时间段（大课）』
		# NUM_COLS = 5;		#列数，代表『一周5天』
		# NUM_MEMBERS = 12;	#组员数量，12人
		# UserName = 'JaunteZhou'		# MAC用户系统用户名，对目录有影响
		# ROOT = '/' + 'Users' + '/' + UserName + '/' + 'Desktop' + '/' + 'Ddd_Pkkb';

	# 某组内组员列表，变量类型『列表』，初始化为空，用于存储12*6个组员姓名
	members_list = []
		#for i in range(0, PARA.NUM_MEMBERS):		#假设成员12人，range从0到12且不包括12
			#members_list = members_list + [""];
		#print "Length of members_list :", len( members_list );

	# 输入第几周，并化为 1 或者 2 ，代表单双周
	WEEK = input("Please input the week No.")
	week = WEEK % 2
	if week == 0:
		week = 2

	# 空课表类的实例化对象
	empty_class_table = C_EmptyClassTable.EmptyClassTable( week )


	#########################################################
	# # # 依次读取队员空课表数据，存储在变量 members_list 中 # # #
	#########################################################
	for g in range(1, PARA.NUM_GROUP + 1) :
		# 在成员表 members_list 中插入一个数组用于存储新的成员变量
		members_list += [ [] ]

		# Excel工作簿名，依次遍历6个现场组组员空课时间表
		excel_name = HEAD.Get_XLS_Full_Name( PARA.DATA_DIR, str(g) )
		# 返回值为False，则文件不存在，跳过
		if not excel_name :
			continue

		# 返回文件名，则打开文件
		excel_bk = xlrd.open_workbook( excel_name )

		for No in range(0, PARA.NUM_MEMBERS) :
			
			#print "No.", No

			# 获取Excel中12个sheet表的表名 # 即组员名
			mbr_name = excel_bk.sheet_names()[ No ]
			# print "第", g, "组 ", "组员列表 :"
			# print sheet_name_list
			mbr_ct = excel_bk.sheet_by_index( No )

			# 新声明一个督导队队员对象
			ddd_mbr = C_DddMember.DddMember(mbr_name, g, g*100+No)
			# 初始化队员对象中的课表
			ddd_mbr.Insert_ClassTable( week, mbr_ct )

			# 保存该队员至队员列表中
			members_list[g-1] += [ ddd_mbr ]

			empty_class_table.Insert_Member( ddd_mbr )
			#members_list[g-1][No].Show_Info()

	#########################################################
	####                     录入完毕                      ###
	#########################################################

	empty_class_table.Show_Info()

	print "输入完毕！开始计算！"

	empty_class_table.Arrange_Class_Table( members_list )


	print "开始写入文件！"
	# Excel工作簿名，依次遍历6个现场组组员空课时间表
	excel_name = "第" + str(WEEK) + "周排课表.xls"

	# 返回文件名，则打开文件
	wb = xlwt.Workbook()
	ws = wb.add_sheet("Check Class Table")
	ws_all = wb.add_sheet("Empty Class Table")

	for day in range(0, PARA.NUM_COLS) :
		for period in range(0, PARA.NUM_ROWS) :
			# 写入“排课表”
			info = ""
			for mbr_id in empty_class_table.arranged_checkclass[day][period]:
				mbr_g = mbr_id / 100 - 1
				mbr_No = mbr_id % 100
				info += members_list[mbr_g][mbr_No].name
				info += "("
				info += str(mbr_g+1)
				info += "), "
			# end for id & group
			ws.write( period+1, day+1, info)

			# 写入“空课总表”
			info = ""
			for mbr_id in empty_class_table.all_emptyclass[day][period]:
				mbr_g = mbr_id / 100 - 1
				mbr_No = mbr_id % 100
				info += members_list[mbr_g][mbr_No].name
				info += "("
				info += str(mbr_g+1)
				info += "), "
			# end for id & group
			ws_all.write( period+1, day+1, info)
		# end for period
	# end for day

	
	ws.write(0,1,u"周一")
	ws.write(0,2,u"周二")
	ws.write(0,3,u"周三")
	ws.write(0,4,u"周四")
	ws.write(0,5,u"周五")
	
	ws.write(1,0,u"12节")
	ws.write(2,0,u"34节")
	ws.write(3,0,u"56节")
	ws.write(4,0,u"78节")

	ws_all.write(0,1,u"周一")
	ws_all.write(0,2,u"周二")
	ws_all.write(0,3,u"周三")
	ws_all.write(0,4,u"周四")
	ws_all.write(0,5,u"周五")

	ws_all.write(1,0,u"12节")
	ws_all.write(2,0,u"34节")
	ws_all.write(3,0,u"56节")
	ws_all.write(4,0,u"78节")
	

	root = PARA.DATA_DIR + "/" + excel_name
	wb.save(root)

	print "写入完毕！"
	print "文件路径为", root
#################
### 主函数结束 ###
#################

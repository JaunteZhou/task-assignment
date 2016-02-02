#-*- coding: utf8 -*-

# Python 2.7.6
# Author ： JaunteZhou（周潮）
# Time ： 2016.02.01
# Type : V1.0.20160201
# File : Class EmptyClassTable		空课表类文件

import xlrd
import HEAD
#import C_DddMember

PARA = HEAD.Parameter();

class EmptyClassTable :
	def __init__(self, week, filename = 'CCT_Sum') :
		self.week = week

		# 完整组员空课表
		self.all_emptyclass = []

		# 已安排查课组员空课表
		self.arranged_checkclass = []

		# 查课表总数
		self.sum_checkclass_table = []
		
		# 每个时间段可安排查课人数，即空课人数			
		self.num_emptyclass_table = []

		# 打开查课表总数表
		excel_name = HEAD.Get_XLS_Full_Name( PARA.DATA_DIR, filename )
		# 返回值为False，则文件不存在，跳过
		if not excel_name :
			return False

		# 返回文件名，则打开文件
		excel_bk = xlrd.open_workbook( excel_name )

		# 打开指定第几个表格
		# 所需排课的总数
		sheet1 = excel_bk.sheet_by_index( week * 2 - 2 )
		# 特殊的排课要求
		sheet2 = excel_bk.sheet_by_index( week * 2 - 1 )

		# 初始化各个列表
		for i in range(0, PARA.NUM_COLS) :
			self.all_emptyclass += [[]]
			self.arranged_checkclass += [[]]
			self.sum_checkclass_table += [[]]
			self.num_emptyclass_table += [[]]
			for j in range(0, PARA.NUM_ROWS) :
				self.all_emptyclass[i] += [ {} ]		# 最内层为字典类型
				self.arranged_checkclass[i] += [ {} ]
				self.sum_checkclass_table[i] += [ sheet1.cell( j, i ).value ]
				self.num_emptyclass_table[i] += [0]
		# 结束 初始化各个列表 
	######################################################
	######################################################
	######################################################

	#####################
	# # # 排课主程序 # # #
	#####################
	def Arrange_Class_Table(self, mbr_list):
		P = input('input the period No.(0~9)')

		p_day = P / 2
		p_half = P % 2

		print ""
		for day in range(0, PARA.NUM_COLS) :
			for period in range(0, PARA.NUM_ROWS) :
				if p_day == day and p_half == period / 2:
					print "跳过第", day+1, "天第", period+1, "时间段！！！"
					continue

				print ""

				# 判断是否为空
				if self.sum_checkclass_table[day][period] == '' :
					print "第", day+1, "天第", period+1, "时间段为空"
					continue
				elif self.sum_checkclass_table[day][period] == 0 :
					print "第", day+1, "天第", period+1, "时间段为空"
					continue
				# end if

				NUM_NEED = int(self.sum_checkclass_table[day][period])
				for num in range(0, NUM_NEED ) :
					print "开始安排第", day+1, "天第", period+1, "时间段第", num+1, "个队员"
					# 找到一个空闲队员，并将其设置为不空闲
					mbr_id = self.Find_Free_Member(mbr_list, day, period)

					if mbr_id == None :
						deep = 0
						while mbr_id == None:
							if deep == 4 :
								print "交换第", day+1, "天第", period+1, "时间段的队员 深度为超过4，交换失败！"
								break
							# end if
							deep += 1
							print "开始尝试交换第", day+1, "天第", period+1, "时间段的队员 深度为", deep
							mbr_id = self.Exchange_Free_Member(mbr_list, day, period, deep)
						# end while
					# end if
					print "ID 为", mbr_id
				# end for num
			# end for period
		# end for day

		day = p_day
		for period in range(p_half*2, p_half*2+2) :
			print ""

			# 判断是否为空
			if self.sum_checkclass_table[day][period] == '' :
				print "第", day+1, "天第", period+1, "时间段为空"
				continue
			elif self.sum_checkclass_table[day][period] == 0 :
				print "第", day+1, "天第", period+1, "时间段为空"
				continue
			# end if

			NUM_NEED = int(self.sum_checkclass_table[day][period])
			for num in range(0, NUM_NEED ) :
				print "开始安排第", day+1, "天第", period+1, "时间段第", num+1, "个队员"
				# 找到一个空闲队员，并将其设置为不空闲
				mbr_id = self.Find_Free_Member(mbr_list, day, period)
				
				if mbr_id == None :
					deep = 0
					while mbr_id == None:
						if deep == 4 :
							print "交换第", day+1, "天第", period+1, "时间段的队员 深度为超过4，交换失败！"
							break
						# end if
						deep += 1
						print "开始尝试交换第", day+1, "天第", period+1, "时间段的队员 深度为", deep
						mbr_id = self.Exchange_Free_Member(mbr_list, day, period, deep)
					# end while
				# end if
				print "ID 为", mbr_id
			# end for num
		# end for period

		self.Show_Info()
		return True
	#####################
	#####################
	#####################

	########################################
	# # # 找一个该日该时间段一个空闲的队员 # # #
	########################################
	def Find_Free_Member(self, m_list, day, period):
		for m_id in self.all_emptyclass[day][period] :
			m_g = m_id / 100 - 1
			m_No = m_id % 100
			if m_list[m_g][m_No].sum_checkclass == 0 :
				m_list[m_g][m_No].sum_checkclass += 1
				m_list[m_g][m_No].day_checkclass = day
				m_list[m_g][m_No].period_checkclass = period
				self.Insert_Arranged_EmptyClassTable(m_id, day, period)
				return m_id
			# end if
		# end for
		return None
	########################################
	########################################
	########################################

	###############################################
	# # # 找一个该日该时间段一个空闲的队员用于交换 # # #
	###############################################
	def Exchange_Free_Member(self, m_list, day, period, deep):
		for m_id in self.all_emptyclass[day][period] :
			m_g = m_id / 100 - 1
			m_No = m_id % 100
			ex_day = m_list[m_g][m_No].day_checkclass
			ex_period = m_list[m_g][m_No].period_checkclass

			if deep == 1 :
				ex_id = self.Find_Free_Member( m_list, ex_day, ex_period)
				if ex_id == None :
					continue
				# end if
			elif deep > 1 :
				ex_id = self.Exchange_Free_Member( m_list, ex_day, ex_period, deep-1)
				if ex_id == None :
					continue
				# end if
			# end if

			# print "交换成功"

			# 将 m_id 用户从 TA 原有的时间段删除
			self.Delete_Arranged_EmptyClassTable(m_id, ex_day, ex_period)
			# 加入到现在所要求的这个时间
			m_list[m_g][m_No].day_checkclass = day
			m_list[m_g][m_No].period_checkclass = period
			self.Insert_Arranged_EmptyClassTable(m_id, day, period)
			return m_id
		# end for
		return None
	###############################################
	###############################################
	###############################################
	

	def Insert_Arranged_EmptyClassTable(self, m_id, day, period) :
		m_g = m_id / 100
		self.arranged_checkclass[day][period][m_id] = m_g
		print "成功插入第", day+1, "天第", period+1, "时间段 编号为", m_id, "的队员"
		return True


	def Delete_Arranged_EmptyClassTable(self, m_id, day, period) :
		self.arranged_checkclass[day][period].pop( m_id )
		print "成功删除第", day+1, "天第", period+1, "时间段 编号为", m_id, "的队员"
		# self.Show_ECT_Arranged()
		return True


	###################################
	# # # 将队员信息插入完整空课表中 # # #
	###################################
	def Insert_Member(self, ddd_member) :
		for day in range(0, PARA.NUM_COLS) :
			for period in range(0, PARA.NUM_ROWS) :
				self.__insert_to_all_emptyclass_table( ddd_member, day, period)
			# end for period
		# end for day
	###################################
	###################################
	###################################

	##########################################################
	# # # 内部函数：具体分3种情况，将队员信息插入到完整空课表中 # # #
	##########################################################
	def __insert_to_all_emptyclass_table(self, dmbr, day, period) :
		if ( dmbr.class_table[ day ][ period ] == self.week ) :		# week 值为1或2
			# 代表本周有空
			self.all_emptyclass[ day ][ period ][ dmbr.id ] = dmbr.group
			self.num_emptyclass_table[ day ][ period ] += 1
		elif ( dmbr.class_table[ day ][ period ] == 3 ) :
			# "3"代表单双周都有空
			self.all_emptyclass[ day ][ period ][ dmbr.id ] = dmbr.group
			self.num_emptyclass_table[ day ][ period ] += 1
		#else if ( "0" ):
			# "0"代表单双周都无空
	##########################################################
	##########################################################
	##########################################################

	def Show_All_EmptyClassTable(self) :
		print "本周原始总空课表 :", self.all_emptyclass
		print ""

	def Show_Arranged_CheckClassTable(self) :
		print "本周已安排查课表 :", self.arranged_checkclass
		print ""

	def Show_Sum_CheckClassTable(self) :
		print "本周总课表数 :", self.sum_checkclass_table
		print ""

	def Show_Num_EmptyClassTable(self) :
		print "本周可安排人数表 :", self.num_emptyclass_table
		print ""

	def Show_Info(self) :
		self.Show_All_EmptyClassTable()
		self.Show_Sum_CheckClassTable()
		self.Show_Num_EmptyClassTable()
		self.Show_Arranged_CheckClassTable()

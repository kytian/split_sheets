# -*- coding: utf-8 -*-
# @Author: tiankaiyun
# @Date:   2017-08-21 00:42:57
# @Last Modified by:   tiankaiyun
# @Last Modified time: 2017-08-22 02:28:44

import os
import sys
import configparser
import win32com.client
from win32com.client import Dispatch
import pandas as pd

def getAddr(addr_str):
	cell_num = len(addr_str.split(':'))
	leftLetter = ''
	leftNum = 0
	rightLetter = ''
	rightNum = 0
	if cell_num == 1:
		leftLetter = addr_str.split('$')[1]
		leftNum = int(addr_str.split('$')[2])
		rightLetter = leftLetter
		rightNum = leftNum
	else:
		leftLetter = addr_str.split(':')[0].split('$')[1]
		leftNum = int(addr_str.split(':')[0].split('$')[2])
		rightLetter = addr_str.split(':')[1].split('$')[1]
		rightNum = int(addr_str.split(':')[1].split('$')[2])
	return cell_num, leftLetter, leftNum, rightLetter, rightNum

rownum = 405

if __name__ == '__main__':
	in_file = '2018.xlsx'
	# read first col with pandas
	# firstColData = pd.read_excel(in_file, sheet_name = 0, usecols = 'A')
	# print(len(firstColData))
	# gongdiList = {}
	# for each in firstColData:
	# 	print(each)
	# 	if each not in gongdiList:
	# 		gongdiList[each] = 1
	# 	else:
	# 		gongdiList[each] = gongdiList[each] + 1
	# for each in gongdiList:
	# 	print(each)
	# print(len(gongdiList))
	# sys.exit(0)

	app = win32com.client.Dispatch('Excel.Application')
	app.Visible = 1
	in_file = os.path.join(os.getcwd(), '2018.xlsx')
	inbook = app.Workbooks.Open(in_file)
	out_file = os.path.join(os.getcwd(), 'out2018.xlsx')
	outbook = app.Workbooks.Add()
	tmpCellNum, tmpLeftLetter, tmpLeftNum, tmpRightLetter, tmpRightNum = getAddr(inbook.sheets[0].UsedRange.Address)
	print(tmpCellNum)
	print(tmpLeftLetter)
	print(tmpLeftNum)
	print(tmpRightLetter)
	print(tmpRightNum)
	print(tmpLeftLetter+str(tmpLeftNum+1)+':'+tmpLeftLetter+str(tmpRightNum))
	# tmpdata = inbook.sheets[0].Range(tmpLeftLetter+str(tmpLeftNum)+':'+tmpLeftLetter+str(tmpRightNum)).value
	tmpdata = inbook.sheets[0].Range('A2:A' + str(rownum)).value
	print(len(tmpdata))
	gongdiAllData = []
	gongdiOrderList = []
	for each in tmpdata:
		gongdiAllData.append(each[0])
		if each[0] not in gongdiOrderList:
			gongdiOrderList.append(each[0])
	print(gongdiAllData[0])
	print(len(gongdiAllData))
	print(gongdiOrderList)
	# new_sheet = outbook.sheets.Add()
	# new_sheet.Name = '123456'
	# sys.exit(0)
	gongdiList = {}
	for each in gongdiAllData:
		if each in gongdiList:
			gongdiList[each] = gongdiList[each] + 1
		else:
			gongdiList[each] = 1
	cur_index = rownum
	for i in range(len(gongdiOrderList)):
		print(gongdiOrderList[i])
	# sys.exit(0)
	for i in range(len(gongdiOrderList)):
		each = gongdiOrderList[len(gongdiOrderList) - 1 - i]
	# for each in gongdiList:
		print(each + ':' + str(gongdiList[each]))
		tmpsheet = outbook.sheets.Add()
		tmpsheet.Name = each
		dstRange = tmpsheet.Range('A2')
		inbook.sheets[0].Range('B'+str(cur_index - gongdiList[each] + 1)+':BP'+str(cur_index)).Copy(dstRange)
		cur_index = cur_index - gongdiList[each]
	print(len(gongdiList))
	sys.exit(0)

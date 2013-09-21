#!/usr/bin/python
#coding=utf-8

__author__ = 'justin.bto@gmail.com (Justin Zhou)'

import os
import importFromExcel
import Tkinter

from adspygoogle import DfaClient
from tkFileDialog import askopenfilename

#from adspygoogle.dfa.DfaErrors import DfaApiError
#from adspygoogle.common import GenericApiService

HOME = os.path.expanduser('~')

#testFilePath = 'C:\Users\Justin.zhou\Desktop\excelContainsFloodlight.xlsx'

def create_spotlight_group(spotlight_service, spotlight_group_name):
	
	#判断是否已经存在该SpotlightGroup
	result = get_spotlightGroup(spotlight_service, spotlight_group_name)
	
	#如果不存在，则新建该SpotlightGroup
	if not result:
		
		print 'SpotlightGroup \'%s\' is not existed, so we will create the Spotlight Group first...' % spotlight_group_name
		
		# 构建Spotlight group需要的参数
		spotlight_activity_group = {
			'name': spotlight_group_name,
			'groupType': '1',
			'spotlightConfigurationId' : '2218289'
		}
		
		new_spotlight_group_id = spotlight_service.saveSpotlightActivityGroup(spotlight_activity_group)['id']
		return new_spotlight_group_id
	
	#如果存在，则返回该SpotlightGroup的ID
	else:
		print 'SpotlightGroup \'%s\' is existed, so we will create the spotlight activity directly...' % spotlight_group_name
		return result['id']

def create_spotlight(spotlight_service, spotlight_group_id, spotlight_name, spotlight_type_id, tag_method_type_id, expected_url, image_enabled):
	# 构建Spotlight需要的参数
	spotlight_activity = {
		'name': spotlight_name,
		'activityGroupId': spotlight_group_id,
		'activityTypeId': '%d' % spotlight_type_id,
		'tagMethodTypeId': '%d' % tag_method_type_id,
		'expectedUrl': expected_url,
		'imageTagsEnabled': image_enabled
	}
	
	try:
		new_spotlight = spotlight_service.SaveSpotlightActivity(spotlight_activity)[0]
		
		# 打印出程序运行结果
		print 'Spotlight activity \'%s\' with ID \'%s\' was created in \'%s\'.' %  new_spotlight['name'], new_spotlight['id'], new_spotlight['activityGroupId']
	
	#如果出现异常，则提示新建Spotlight错误
	except:
		print 'Failed in creating Spotlight activity \'%s\'' % spotlight_name
	
	print '----------------------------------------------------------------------------'

def get_spotlightGroup(spotlight_service, search_string):
	
	search_criteria = {
		'advertiserId': '2218289',
		'searchString': search_string,
		'type': '1'
	}
	
	results = spotlight_service.getSpotlightActivityGroups(search_criteria)[0]
	if results['records']:
		return results['records'][0]
	else:
		return False


def main(client):

        #弹出窗口让用户选择需要导入的excel文件
	root = Tkinter.Tk()
	file_path = askopenfilename(title = '选择需要导入的excel文件：')
	root.destroy()
	spotlight_dic = importFromExcel.read_spotlight_dic(file_path)
	
	# 初始化Service对象
	spotlight_service = client.GetSpotlightService(version='v1.20')
	
	# 根据导入的表格创建spotlight
	for i in range(len(spotlight_dic['spotlight_group_name'])):
		
		#分别获取spotlight_dict中每行的各参数值
		spotlight_group_name = spotlight_dic['spotlight_group_name'][i]
		spotlight_name = spotlight_dic['spotlight_name'][i]
		spotlight_type_id = spotlight_dic['spotlight_type_id'][i]
		tag_method_type_id = spotlight_dic['tag_method_type_id'][i]
		expected_url = spotlight_dic['expected_url'][i]
		image_enabled = spotlight_dic['image_enabled'][i]
		
		#首先创建Spotlight Group 并返回ID
		spotlight_group_id = create_spotlight_group(spotlight_service, spotlight_group_name)
		create_spotlight(spotlight_service, spotlight_group_id, spotlight_name, spotlight_type_id, tag_method_type_id, expected_url, image_enabled)
	print 'The import task has finished!'

if __name__ == '__main__':
	# 初始化client对象
	client = DfaClient(path=HOME)
	main(client)
	os.system('pause')

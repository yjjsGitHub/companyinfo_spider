#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2021/4/3 9:03
# @Author  : imyjj

import os
import time
# dev prod
env = 'prod'
file_suffix = '.xls'
input_file_name = '\\企业信息' + file_suffix
output_file_name = '\\输出清单%s' + file_suffix
project_curPath = os.path.abspath(os.path.dirname(__file__))
# project_root_Path = project_curPath[:project_curPath.find("qqc_spider\\")+len("qqc_spider\\")]  # 获取myProject，也就是项目的根路径
project_same_level_Path = project_curPath[:project_curPath.find("IdeaProjects\\")]  # 获取myProject，也就是项目的同级根路径

# phantomjs_path = os.path.abspath(project_root_Path + r'\plugin\phantomjs.exe') # 获取phantomjs.exe文件的路径
phantomjs_name = '\\plugin\\phantomjs.exe'
# phantomjs_path = curPath + phantomjs_name  # 获取phantomjs.exe文件的路径

# excel_input_path = curPath + input_file_name  # 获取excel文件的路径
# excel_output_path = curPath + output_file_name  # 获取excel文件的路径
excel_out_put_titles = {'序号':0, '机构名称':1, '原地址':2, '备注':3, '法定代表人':4, '注册资本':5, '成立日期':6, '邮箱':7, '官网':8, '地址':9, '业务范围':10, '电话':11, '登记状态':12, '统一社会信用代码':13, '社会组织类型':14, '证书有效期':15, '同电话企业':16, '爱企查':17}


verison = time.strftime('%Y%m%d%H%M%S', time.localtime(time.time()))



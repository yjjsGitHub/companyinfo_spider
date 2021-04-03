#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2021/4/3 9:03
# @Author  : imyjj

import os
file_name = '输出清单.xls'
curPath = os.path.abspath(os.path.dirname(__file__))
project_root_Path = curPath[:curPath.find("qqc_spider\\")+len("qqc_spider\\")]  # 获取myProject，也就是项目的根路径

# phantomjs_path = os.path.abspath(project_root_Path + r'\plugin\phantomjs.exe') # 获取phantomjs.exe文件的路径
phantomjs_path = os.path.join(os.path.expanduser("~"), 'Desktop') + r'\phantomjs.exe' # 获取phantomjs.exe文件的路径

excel_path = os.path.join(os.path.expanduser("~"), 'Desktop') + '\\' + file_name # 获取excel文件的路径

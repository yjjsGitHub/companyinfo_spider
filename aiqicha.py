#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2021/3/23 22:48
# @Author  : imyjj

from selenium import webdriver
import requests
import re
import json
import os
import win32com
from win32com.client import Dispatch, constants
import pythoncom
import xlrd
import xlwt

import traceback
from config.config import *

headers = {'User-Agent': 'Chrome/76.0.3809.132'}

#需要安装phantomjs，然后将phantomjs.exe路径指定到path
# path = r'D:\SoftWares\phantomjs-2.1.1-windows\bin\phantomjs.exe'
# print(path)
# 正则表达式提取数据
re_get_js = re.compile(r'<script>([\s\S]*?)</script>')
re_resultList = re.compile(r'"resultList":(\[{.+?}\]),"totalNumFound')
has_statement = False


def get_company_info(sheet, name, col_index_dict, row_nom, curPath):
    '''
        @func: 通过百度企业信用查询企业基本信息
    '''
    url = 'https://aiqicha.baidu.com/s?q=%s' % name
    res = requests.get(url, headers=headers)
    if res.status_code == 200:
        html = res.text
        js = re_get_js.findall(html)[1]
        data = re_resultList.search(js)
        if not data:
            return
        company = json.loads(data.group(1))[0]
        url = 'https://aiqicha.baidu.com/company_detail_{}'.format(company['pid'])

        # 调用环境变量指定的PhantomJS浏览器创建浏览器对象
        driver = webdriver.PhantomJS(curPath + phantomjs_name)
        driver.set_window_size(1366, 768)
        driver.get(url)

        # 获取页面名为wraper的id标签的文本内容
        data = driver.find_element_by_class_name('content-info').text
        data = re.split(r'[：\r\n]', data)
        data_return = {}
        # print(data)
        need_info = ['电话']
        for epoch, item in enumerate(data):
            if item in need_info:
                sheet.write(row_nom, col_index_dict['爱企查'], label=data[epoch + 1].replace('隐藏', ''))
                # print('电话：' + data[epoch + 1].replace('隐藏', ''))
    else:
        print('无法获取%s的企业信息' % name)
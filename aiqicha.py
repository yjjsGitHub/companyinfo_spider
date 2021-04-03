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

import traceback
from config.config import *

SPATH = excel_path  # 需处理的excel文件目录
headers = {'User-Agent': 'Chrome/76.0.3809.132'}

#需要安装phantomjs，然后将phantomjs.exe路径指定到path
# path = r'D:\SoftWares\phantomjs-2.1.1-windows\bin\phantomjs.exe'
path = phantomjs_path
# print(path)
# 正则表达式提取数据
re_get_js = re.compile(r'<script>([\s\S]*?)</script>')
re_resultList = re.compile(r'"resultList":(\[{.+?}\]),"totalNumFound')
has_statement = False


def get_company_info(sheet, name, col_index_dict, row_nom):
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
        driver = webdriver.PhantomJS(path)
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
                sheet.Cells(row_nom, col_index_dict['爱企查']).Value = data[epoch + 1].replace('隐藏', '')
                print('电话：' + data[epoch + 1].replace('隐藏', ''))
    else:
        print('无法获取%s的企业信息' % name)


def run_aiqicha():
    app = win32com.client.Dispatch('Excel.Application')

    # 后台运行，不显示，不警告
    app.Visible = 0
    app.DisplayAlerts = 0

    # 创建新的Excel
    # work_book = app.Workbooks.Add()
    # 新建sheet
    # sheet = work_book.Worksheets.Add()

    # 打开已存在表格，注意这里要用绝对路径
    work_book = app.Workbooks.Open(SPATH)
    sheet = work_book.Worksheets('sheet1')

    # 获取单元格信息 第n行n列，不用-1
    # cell01_value = sheet.Cells(5, 2).Value
    # print("cell01的内容为：", cell01_value)
    try:
        # 记下表头的名字和位置
        col_name_and_col_num_dic = {}
        col_num = 1
        if sheet.Cells(1, col_num).Value is not None:
            col_name = sheet.Cells(1, col_num).Value
            while col_name is not None:
                col_name_and_col_num_dic[col_name] = col_num
                col_num = col_num + 1
                col_name = sheet.Cells(1, col_num).Value
        else:
            # 关闭表格
            work_book.Close()
            app.Quit()
            return
        # print(col_name_and_col_num_dic['发证机关'])
        row_num = 2
        col_num = 2

        company_name = sheet.Cells(row_num, col_num).Value
        while company_name is not None:
            try:
                if sheet.Cells(row_num, col_name_and_col_num_dic['同电话企业']).Value is not None and \
                        sheet.Cells(row_num, col_name_and_col_num_dic['爱企查']).Value is None:
                    print('企业名称：' + company_name)
                    get_company_info(sheet, company_name, col_name_and_col_num_dic, row_num)
            except Exception as e:
                print(company_name + '失败', e)
            # 保存表格
            work_book.Save()
            row_num = row_num + 1
            company_name = sheet.Cells(row_num, col_num).Value
    except Exception as ex:
        traceback.print_exc(ex)
        print('失败了')

    # 写入表格信息
    # sheet.Cells(2, 1).Value = "win32com"



    # 另存为实现拷贝
    # work_book.SaveAs(getScriptPath() + "\\new.xlsx")

    # 关闭表格
    work_book.Close()
    app.Quit()


if __name__ == '__main__':
    run_aiqicha()
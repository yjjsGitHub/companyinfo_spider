#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2021/3/23 22:48
# @Author  : imyjj

import requests
import random
from bs4 import BeautifulSoup

import win32com
from win32com.client import Dispatch, constants
import pythoncom

import traceback

import _thread
import time
from config.config import *

SPATH = excel_path  # 需处理的excel文件目录

url = r'https://www.qcc.com/web/search?key='
url_aiqicha = r'https://aiqicha.baidu.com/s?t=0&q='

# 用于模拟http头的User-agent
ua_list = [
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10.6; rv2.0.1) Gecko/20100101 Firefox/4.0.1",
    "Mozilla/5.0 (Windows NT 6.1; rv2.0.1) Gecko/20100101 Firefox/4.0.1",
    "Opera/9.80 (Macintosh; Intel Mac OS X 10.6.8; U; en) Presto/2.8.131 Version/11.11",
    "Opera/9.80 (Windows NT 6.1; U; en) Presto/2.8.131 Version/11.11",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_7_0) AppleWebKit/535.11 (KHTML, like Gecko) Chrome/17.0.963.56 Safari/535.11"
]

user_agent = random.choice(ua_list)

# 登录的cookie
log_cookie = r'UM_distinctid=1785f64245d53-0f746198dcd7c4-5771031-151800-1785f64245e31c; zg_did=%7B%22did%22%3A%20%221785f64248480-0a9724a8aca673-5771031-151800-1785f642486104%22%7D; _uab_collina=161650810290177689079926; acw_tc=b7f0d73116171926370131817edab42c2ab780a6ac68b3c03205778cde; QCCSESSID=qcggnsbe3k4ja76ljcuvjc8mj7; CNZZDATA1254842228=192233519-1616507339-https%253A%252F%252Fwww.baidu.com%252F%7C1617188737; hasShow=1; zg_de1d1a35bfa24ce29bbf2c7eb17e6c4f=%7B%22sid%22%3A%201617192638547%2C%22updated%22%3A%201617193276064%2C%22info%22%3A%201617192638552%2C%22superProperty%22%3A%20%22%7B%5C%22%E5%BA%94%E7%94%A8%E5%90%8D%E7%A7%B0%5C%22%3A%20%5C%22%E4%BC%81%E6%9F%A5%E6%9F%A5%E7%BD%91%E7%AB%99%5C%22%7D%22%2C%22platform%22%3A%20%22%7B%7D%22%2C%22utm%22%3A%20%22%7B%7D%22%2C%22referrerDomain%22%3A%20%22www.baidu.com%22%2C%22cuid%22%3A%20%2237284da7907f116e869a3b5495d0b8ad%22%2C%22zs%22%3A%200%2C%22sc%22%3A%200%7D'
# 不登录的cookie
unlog_cookie = r'UM_distinctid=1785f64245d53-0f746198dcd7c4-5771031-151800-1785f64245e31c; zg_did=%7B%22did%22%3A%20%221785f64248480-0a9724a8aca673-5771031-151800-1785f642486104%22%7D; _uab_collina=161650810290177689079926; QCCSESSID=qcggnsbe3k4ja76ljcuvjc8mj7; hasShow=1; acw_tc=b7f0d73716171954630282654e955346332ad16abe402431496d1c455d; CNZZDATA1254842228=192233519-1616507339-https%253A%252F%252Fwww.baidu.com%252F%7C1617194137; zg_de1d1a35bfa24ce29bbf2c7eb17e6c4f=%7B%22sid%22%3A%201617192638547%2C%22updated%22%3A%201617195464637%2C%22info%22%3A%201617192638552%2C%22superProperty%22%3A%20%22%7B%7D%22%2C%22platform%22%3A%20%22%7B%7D%22%2C%22utm%22%3A%20%22%7B%7D%22%2C%22referrerDomain%22%3A%20%22www.baidu.com%22%2C%22cuid%22%3A%20%2237284da7907f116e869a3b5495d0b8ad%22%2C%22zs%22%3A%200%2C%22sc%22%3A%200%7D'

headers = {
    'User-Agent': user_agent,
    'cookie': unlog_cookie
}


#
def findCompanyDetailHtml(name):
    # 向指定的url地址发送请求，并返回服务器响应的类文件对象
    response = requests.get(url=url + name, headers=headers)

    # 服务器返回的类文件对象支持python文件对象的操作方法
    # print(response.text)
    soup = BeautifulSoup(response.text, 'lxml')
    a_list = soup.find_all('a')
    # print(a_list)
    # print('###')
    for a in a_list:
        # print(a)
        # print('##')
        stag = str(a)
        if 'https://www.qcc.com/firm/' in stag:
            uri = str(a.get('href'))
            if uri.endswith('html'):
                detail_response = requests.get(uri, headers=headers)
                # print(detail_response.text)
                detail_soup = BeautifulSoup(detail_response.text, 'lxml')
                cominfo = detail_soup.find('section', id='Cominfo')
                # print(cominfo)
                cominfo_list = cominfo.find_all('td')
                for i, td in enumerate(cominfo.find_all('td')):
                    if td['class'] == ['tb']:
                        print(td.text, cominfo_list[i + 1].text.strip())
                        print('##')


#
def findCompanyDetailHtml(sheet, name, col_index_dict, row_nom):
    # 向指定的url地址发送请求，并返回服务器响应的类文件对象
    response = requests.get(url=url + name, headers=headers)

    # 服务器返回的类文件对象支持python文件对象的操作方法
    # print(response.status_code)
    soup = BeautifulSoup(response.text, 'lxml')
    # 获取table
    table = soup.find('table', class_='ntable ntable-list')
    tr_list = table.find_all('tr')
    not_find_flag = True
    for tr in tr_list:
        if not not_find_flag:
            return
        a = tr.find('a')
        # print(a_list)
        # print('###')

        # print(a)
        # print('##')
        stag = str(a).replace('<em>','').replace('</em>', '')
        if 'https://www.qcc.com/firm/' in stag and name in stag:
            not_find_flag = False
            uri = str(a.get('href'))
            if uri.endswith('html'):
                detail_response = requests.get(uri, headers=headers)
                # print(detail_response.text)
                detail_soup = BeautifulSoup(detail_response.text, 'lxml')
                record1(col_index_dict, detail_soup, row_nom, sheet)

                record2(col_index_dict, detail_soup, row_nom, sheet)

                # 备注成功
                sheet.Cells(row_nom, col_index_dict['备注']).Value = 1
                print('查询成功')
    if not_find_flag:
        sheet.Cells(row_nom, col_index_dict['备注']).Value = '查不到此公司信息'
        print('查不到此公司信息')


def record1(col_index_dict, detail_soup, row_nom, sheet):
    div = detail_soup.find('div', class_='dcontent')
    for row in div.find_all('div', class_='row'):
        span_list = row.find_all('span')
        for i, span in enumerate(span_list):
            text = span.text
            if span.get('class') == ['fc']:
                sub_span_list = span.find_all('span')
                if col_index_dict.get(sub_span_list[0].text[0:-1]) is not None:
                    value = span_list[i + 1].text.strip()
                    sheet.Cells(row_nom, col_index_dict[sub_span_list[0].text[0:-1]]).Value = value
            elif span.get('class') == ['cdes'] and col_index_dict.get(text[0:-1]) is not None:
                value = span_list[i + 1].text.strip()
                if '同电话' in value:
                    sheet.Cells(row_nom, col_index_dict['同电话企业']).Value = 1
                sheet.Cells(row_nom, col_index_dict[text[0:-1]]).Value = value


def record2(col_index_dict, detail_soup, row_nom, sheet):
    cominfo = detail_soup.find('section', id='Cominfo')
    # print(cominfo)
    cominfo_list = cominfo.find_all('td')
    for i, td in enumerate(cominfo.find_all('td')):
        text = td.text.replace('""', '').strip()
        if td.get('class') == ['tb'] and col_index_dict.get(text) is not None:
            next_td = cominfo_list[i + 1]
            if next_td.get('class') == ['boss-td']:
                value = next_td.find('div', class_='clearfix').find('div', class_='bpen').find('a', class_='bname')\
                    .find('h2').text
            else:
                value = cominfo_list[i + 1].text.strip()
            # print(text, value)
            sheet.Cells(row_nom, col_index_dict[text]).Value = value


def run_qiqicha():
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
                if sheet.Cells(row_num, col_name_and_col_num_dic['备注']).Value is None:
                    print('企业名称：' + company_name)
                    findCompanyDetailHtml(sheet, company_name, col_name_and_col_num_dic, row_num)
            except Exception as e:
                print(company_name + '失败', e)
                # 备注成功
                sheet.Cells(row_num, col_name_and_col_num_dic['备注']).Value = str(e)
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
    # findCompanyDetailHtml('广州市天河区三立教育培训中心')
    # findCompanyDetailHtml('广州市天河区高越舞蹈培训中心')
    run_qiqicha()
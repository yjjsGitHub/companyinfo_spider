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
import xlrd
import xlwt

import traceback

import _thread
import time
from config.config import *
import aiqicha

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

# 不登录的cookie
unlog_cookie = r'UM_distinctid=1785f64245d53-0f746198dcd7c4-5771031-151800-1785f64245e31c; zg_did=%7B%22did%22%3A%20%221785f64248480-0a9724a8aca673-5771031-151800-1785f642486104%22%7D; _uab_collina=161650810290177689079926; QCCSESSID=qcggnsbe3k4ja76ljcuvjc8mj7; hasShow=1; acw_tc=b7f0d73716171954630282654e955346332ad16abe402431496d1c455d; CNZZDATA1254842228=192233519-1616507339-https%253A%252F%252Fwww.baidu.com%252F%7C1617194137; zg_de1d1a35bfa24ce29bbf2c7eb17e6c4f=%7B%22sid%22%3A%201617192638547%2C%22updated%22%3A%201617195464637%2C%22info%22%3A%201617192638552%2C%22superProperty%22%3A%20%22%7B%7D%22%2C%22platform%22%3A%20%22%7B%7D%22%2C%22utm%22%3A%20%22%7B%7D%22%2C%22referrerDomain%22%3A%20%22www.baidu.com%22%2C%22cuid%22%3A%20%2237284da7907f116e869a3b5495d0b8ad%22%2C%22zs%22%3A%200%2C%22sc%22%3A%200%7D'

headers = {
    'User-Agent': user_agent,
    'cookie': unlog_cookie
}


def findCompanyDetailHtml(sheet, name, col_index_dict, row_nom, curPath):
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
        stag = str(a).replace('<em>', '').replace('</em>', '')
        if 'https://www.qcc.com/firm/' in stag and name in stag:
            not_find_flag = False
            uri = str(a.get('href'))
            if uri.endswith('html'):
                detail_response = requests.get(uri, headers=headers)
                # print(detail_response.text)
                detail_soup = BeautifulSoup(detail_response.text, 'lxml')
                record1(col_index_dict,name, detail_soup, row_nom, sheet, curPath)

                record2(col_index_dict, detail_soup, row_nom, sheet)

                # 备注成功
                # sheet.Cells(row_nom, col_index_dict['备注']).Value = 1
                sheet.write(row_nom, col_index_dict['备注'], label=1, style=xlwt.XFStyle())
                print('查询成功')
    if not_find_flag:
        sheet.write(row_nom, col_index_dict['备注'], label='查不到此公司信息')
        print('查不到此公司信息')


def record1(col_index_dict, company_name, detail_soup, row_nom, sheet, curPath):
    div = detail_soup.find('div', class_='dcontent')
    for row in div.find_all('div', class_='row'):
        span_list = row.find_all('span')
        for i, span in enumerate(span_list):
            text = span.text
            if span.get('class') == ['fc']:
                sub_span_list = span.find_all('span')
                if col_index_dict.get(sub_span_list[0].text[0:-1]) is not None:
                    value = span_list[i + 1].text.strip()
                    sheet.write(row_nom, col_index_dict[sub_span_list[0].text[0:-1]], label=value, style=xlwt.XFStyle())
            elif span.get('class') == ['cdes'] and col_index_dict.get(text[0:-1]) is not None:
                value = span_list[i + 1].text.strip()
                if '同电话' in value:
                    sheet.write(row_nom, col_index_dict['同电话企业'], label=value, style=xlwt.XFStyle())
                    aiqicha.get_company_info(sheet, company_name, col_index_dict, row_nom, curPath)
                sheet.write(row_nom, col_index_dict[text[0:-1]], label=value, style=xlwt.XFStyle())


def record2(col_index_dict, detail_soup, row_nom, sheet):
    cominfo = detail_soup.find('section', id='Cominfo')
    # print(cominfo)
    cominfo_list = cominfo.find_all('td')
    for i, td in enumerate(cominfo.find_all('td')):
        text = td.text.replace('""', '').strip()
        if td.get('class') == ['tb'] and col_index_dict.get(text) is not None:
            next_td = cominfo_list[i + 1]
            if next_td.get('class') == ['boss-td']:
                value = next_td.find('div', class_='clearfix').find('div', class_='bpen').find('a', class_='bname') \
                    .find('h2').text
            else:
                value = cominfo_list[i + 1].text.strip()
            # print(text, value)
            sheet.write(row_nom, col_index_dict[text], label=value, style=xlwt.XFStyle())


def run_qiqicha(curPath):
    # win32com内容
    # app = win32com.client.Dispatch('Excel.Application')

    # 后台运行，不显示，不警告
    # app.Visible = 0
    # app.DisplayAlerts = 0

    # 创建新的Excel
    # work_book = app.Workbooks.Add()
    # 新建sheet
    # sheet = work_book.Worksheets.Add()

    # 打开已存在表格，注意这里要用绝对路径
    work_book = xlrd.open_workbook(curPath + input_file_name)
    sheet = work_book.sheet_by_name('sheet1')
    # work_book = app.Workbooks.Open(curPath + input_file_name)
    # sheet = work_book.Worksheets('sheet1')

    # 创建一个workbook 设置编码
    new_work_book = xlwt.Workbook(encoding='utf-8')
    # 创建一个worksheet
    new_worksheet = new_work_book.add_sheet('sheet1', cell_overwrite_ok=True)

    # 获取单元格信息 第n行n列，不用-1
    # cell01_value = sheet.Cells(5, 2).Value
    # print("cell01的内容为：", cell01_value)
    try:
        # 记下表头的名字和位置
        col_name_and_col_num_dic = excel_out_put_titles
        for key in col_name_and_col_num_dic:
            new_worksheet.write(0, col_name_and_col_num_dic[key], label=key, style=xlwt.XFStyle())
        col_num = 1
        company_name_list = sheet.col_values(1, 1)
        input_addre_list = sheet.col_values(2, 1)
        # print(company_name_list)
        row_num = 1
        col_num = 0

        for company_name in company_name_list:
            new_worksheet.write(row_num, 0, label=row_num)
            new_worksheet.write(row_num, 1, label=company_name, style=xlwt.XFStyle())
            new_worksheet.write(row_num, 1, label=input_addre_list[row_num-1], style=xlwt.XFStyle())
            try:
                print('企业名称：' + company_name)
                findCompanyDetailHtml(new_worksheet, company_name, col_name_and_col_num_dic, row_num, curPath)
            except Exception as e:
                print(company_name + '失败', e)
                # 备注成功
                # sheet.Cells(row_num, col_name_and_col_num_dic['备注']).Value = str(e)
                new_worksheet.write(row_num, col_name_and_col_num_dic['备注'], label=str(e), style=xlwt.XFStyle())
            # 保存表格
            # work_book.Save()
            row_num = row_num + 1
            # company_name = sheet.Cells(row_num, col_num).Value
    except Exception as ex:
        traceback.print_exc(ex)
        print('失败了')

    # 写入表格信息
    # sheet.Cells(2, 1).Value = "win32com"

    # 另存为实现拷贝
    # work_book.SaveAs(curPath + output_file_name.replace('%s', time.strftime('%Y%m%d%H%M%S',time.localtime(time.time()))))

    # 保存
    new_work_book.save(
        curPath + output_file_name.replace('%s', verison))

    # 关闭表格
    # work_book.Close()
    # app.Quit()


if __name__ == '__main__':
    # findCompanyDetailHtml('广州市天河区三立教育培训中心')
    # findCompanyDetailHtml('广州市天河区高越舞蹈培训中心')
    run_qiqicha()

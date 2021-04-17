#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2021/4/3 9:01
# @Author  : imyjj

import qiqicha
import aiqicha
from config.config import *

if __name__ == '__main__':
    curPath = os.path.abspath(os.path.dirname(__file__))
    if env == 'dev':
        curPath = curPath.split("\\bin")
        curPath = str(curPath[0]) + '\\file'
        print(curPath)
    qiqicha.run_qiqicha(curPath)
    # qiqicha.run_qiqicha()
    # aiqicha.run_aiqicha()
    # print(phantomjs_path)
    # input("please input any key to exit!")

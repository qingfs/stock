# -*- coding: utf-8 -*-
from docx import Document

from docx.enum.text import WD_ALIGN_PARAGRAPH

from docx.oxml.ns import qn # 中文格式

from docx.shared import Pt # 磅数

from docx.shared import Inches # 图片尺寸

from bs4 import BeautifulSoup

from urllib.request import urlopen

import pandas as pd

import numpy as np

import matplotlib.pyplot as plt

import os


# 测试python异常调试
def func1():
    result = 10 / 1
    raise ValueError('invalid value')


fpath = 'F:\\job\\CMDI\\智能分析\\python.txt'

# 读取
def read_and_write():
    path = pathlib.Path(fpath)
    if path.exists():
        with open(fpath, 'a+') as f:
            f.write("hello world！" + '\n')
    else:
        with open(fpath, 'w+') as f:
            f.write("hello world" + '\n')


if __name__ == '__main__':
    url ='http://quotes.money.163.com/f10/zcfzb_600660.html?type=year'
    html = urlopen(url)
    soup = BeautifulSoup(html,'lxml')
    print(soup)
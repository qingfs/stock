# -*- coding: utf-8 -*-
from docx import Document

from docx.enum.text import WD_ALIGN_PARAGRAPH

from docx.oxml.ns import qn # 中文格式

from docx.shared import Pt # 磅数

from docx.shared import Inches # 图片尺寸

from bs4 import BeautifulSoup

from urllib.request import urlopen
import  requests

import pandas as pd

import numpy as np

import matplotlib.pyplot as plt

import os
import chardet


c_name = '福耀玻璃'

stock_code = '600660'

st = 2015

et = 2020



data_list =['zcfzb','lrb','xjllb']

adrees ='http://quotes.money.163.com'

y_list =['报告日期']+[str(m)+'年' for m in list(range(int(st)-1,int(et)+1))][::-1]


def stock():
    url = 'http://quotes.money.163.com/f10/zcfzb_600660.html?type=year'

    html = urlopen(url)

    soup = BeautifulSoup(html,'lxml')

    div =soup.findAll('div',{'class':'inner_box'})

    df = BeautifulSoup(str(div[0]),features='lxml')

    a = df.findAll('a')

    for each in a:
        if each.string == '下载数据':
            new_html = adrees + each.get("href")
            html1 = urlopen(new_html)


            soup1 = BeautifulSoup(html1,'lxml')

            txt = soup1.text.replace('(万元)','').replace('--','0')

            csv = open(r'F:\stock\python_data\python_data\财务分析\数据采集\%s%s.csv'%(stock_code,'zcfzb'),'w',encoding='utf-8') .write(txt)

            data = pd.read_csv(r'F:\stock\python_data\财务分析\数据采集\%s%s.csv'%(stock_code,'zcfzb'))
            # 读取年份。-2是倒数第二个，由文档本身决定；:4是连续前4个；::1是倒序
            list1 = list(range(int(data.columns[-2][:4]), int(data.columns[1][:4]) + 1))[::-1]

            writer = pd.ExcelWriter(r'F:\stock\python_data\%s%s.xlsx'%(stock_code,'zcfzb'))
            data.to_excel(writer,index=False)
            writer.save()
            writer.close()

            p_list = os.listdir(r'F:\stock\python_data')
            for name in p_list:
                if name.endswith('.csv'):
                    os.remove(r'F:\stock\python_data\%s%s.xlsx'%(stock_code,'zcfzb'))

    path_list = os.listdir(r'C:\Users\admin\Desktop\python_data\财务分析\数据采集')

    data = pd.DataFrame()

    for path in path_list:

        fp = r'C:\Users\admin\Desktop\python_data\财务分析\数据采集\%s' % (path)

        dfs = pd.read_excel(fp,None,usecols=y_list)

        keys = dfs.keys()

        for i in keys:

            df1 = dfs[i]

            data = pd.concat([df1,data])

            data = data.fillna(0)

            data['报告日期'] =data['报告日期'].str.strip()

            data = data.drop_duplicates(subset=['报告日期'],keep='first')

    data = data.set_index('报告日期')

    hb = data.drop('%s年'% (int(st)-1),axis=1)




if __name__ == '__main__':
    stock()


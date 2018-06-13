# -*- coding:utf-8 -*-
'''
                   _ooOoo_
                  o8888888o
                  88" . "88
                  (| -_- |)
                  O\  =  /O
               ____/`---'\____
             .'  \\|     |//  `.
            /  \\|||  :  |||//  \
           /  _||||| -:- |||||-  \
           |   | \\\  -  /// |   |
           | \_|  ''\---/''  |   |
           \  .-\__  `-`  ___/-. /
         ___`. .'  /--.--\  `. . __
      ."" '<  `.___\_<|>_/___.'  >'"".
     | | :  `- \`.;`\ _ /`;.`/ - ` : | |
     \  \ `-.   \_ __\ /__ _/   .-` /  /
======`-.____`-.___\_____/___.-`____.-'======
                   `=---='
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
            佛祖保佑       永无BUG
'''
import os
import csv
import xlrd
import xlwt
import xlsxwriter
import time
import scipy.io as sio
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
def ISA_group_ave_det(t,k):
    #详细信息，包括四分位、用户的计数、均值等
    t1 = time.time()

    #所属群组
    a_group = xlrd.open_workbook('sta\\Filmtrust_isa_%d_%d_all.xlsx'%(t,k),'rb')

    workbook = xlsxwriter.Workbook('sta\\Filmtrust_isa_%d_%d_ave_det.xlsx'%(t,k))
    format_r = workbook.add_format({'bg_color' : '#FF5151'})# 红色 均值
    format_g = workbook.add_format({'bg_color' : '#ADADAD'})# 灰色 计数
    format_b = workbook.add_format({'bg_color' : '#99CC99'})# 棕色 方差

    for x in range(k):
        gr = a_group.sheet_by_index(x+1)#通过获取
        nrows = gr.nrows
        ncols = gr.ncols

        # 2. 创建一个工作表sheet对象
        sheet1 = workbook.add_worksheet('%d'%x)#第一个表存储评分所属群组
        # 复制表格内容
        for i in range(nrows):
            for j in range(ncols):
                sheet1.write(i,j,gr.cell(i,j).value)
        # 写抬头
        sheet1.write(0,ncols,'count',format_g)
        sheet1.write(0,ncols+1,'average',format_r)
        sheet1.write(0,ncols+2,'std',format_b)
        sheet1.write(0,ncols+3,'min')
        sheet1.write(0,ncols+4,'25%')
        sheet1.write(0,ncols+5,'75%')
        sheet1.write(0,ncols+6,'max')
        sheet1.write(nrows,0,'count',format_g)
        sheet1.write(nrows+1,0,'average',format_r)
        sheet1.write(nrows+2,0,'std',format_b)
        sheet1.write(nrows+3,0,'min')
        sheet1.write(nrows+4,0,'25%')
        sheet1.write(nrows+5,0,'75%')
        sheet1.write(nrows+6,0,'max')
        # 统计用户信息

        for i in range(nrows):
            list_u = []
            for j in range(ncols):
                if gr.cell(i,j).value:
                    list_u.append(gr.cell(i,j).value)
            if len(list_u) >=1:
                narray = np.array(list_u)
                sheet1.write(i,ncols,len(list_u),format_g)#记录计数
                sheet1.write(i,ncols+1,narray.mean(),format_r)#记录均值
                sheet1.write(i,ncols+2,narray.var(),format_b)#记录方差
                sheet1.write(i,ncols+3,narray.min())
                sheet1.write(i,ncols+4,np.percentile(narray,25))#25%分位数
                sheet1.write(i,ncols+5,np.percentile(narray,75))#75%分位数
                sheet1.write(i,ncols+6,narray.max())

        # 统计item信息
        for j in range(ncols):
            list_i = []
            for i in range(nrows):
                if gr.cell(i,j).value:
                    list_i.append(gr.cell(i,j).value)
            if len(list_i) >= 1:
                narray = np.array(list_i)
                sheet1.write(nrows,j,len(list_i),format_g)#记录计数
                sheet1.write(nrows+1,j,narray.mean(),format_r)#记录均值
                sheet1.write(nrows+2,j,narray.var(),format_b)#记录方差
                sheet1.write(nrows+3,j,narray.min())
                sheet1.write(nrows+4,j,np.percentile(narray,25))#25%分位数
                sheet1.write(nrows+5,j,np.percentile(narray,75))#75%分位数
                sheet1.write(nrows+6,j,narray.max())

    workbook.close()

    t2 = time.time()
    print("群组统计完成，耗时："+str(t2-t1)+"秒。") #反馈结果

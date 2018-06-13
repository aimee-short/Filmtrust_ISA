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
def ISA_group_ave(t,k):
    t1 = time.time()
    #所属群组
    a_group = xlrd.open_workbook('sta\\Filmtrust_isa_%d_%d_all.xlsx'%(t,k),'rb')

    # 每个群组的用户词典
    group_dict = xlrd.open_workbook('sta\\Filmtrust_isa_%d_%d_sta.xlsx'%(t,k),'rb')
    t_dict = group_dict.sheet_by_name('dict')

    # 1. 创建一个Excel文件
    workbook = xlsxwriter.Workbook('sta\\Filmtrust_isa_%d_%d_ave.xlsx'%(t,k))
    sheet1 = workbook.add_worksheet('ave')#存储群组均值 第一行：item_id 第二行：计数 第三行：均值 第四行：方差

    format_r = workbook.add_format({'bg_color' : '#FF5151'})# 红色 均值
    format_g = workbook.add_format({'bg_color' : '#ADADAD'})# 灰色 计数
    format_b = workbook.add_format({'bg_color' : '#99CC99'})# 绿色 方差

    x = 0# x代表第几组
    for gr in a_group.sheets():
        nrows = gr.nrows
        ncols = gr.ncols

        sheet1.write(x*4,0,'G_%d'%(x+1),format_r)
        sheet1.write(x*4+1,0,'count',format_g)
        sheet1.write(x*4+2,0,'average',format_g)
        sheet1.write(x*4+3,0,'std',format_g)
        # 统计item信息
        for j in range(ncols):
            sheet1.write(x*4,j+1,t_dict.cell(x*3+2,j).value,format_b)# 写入dict  item编号
            list_i = []#有效评分
            for i in range(nrows):
                if gr.cell(i,j).value:
                    list_i.append(gr.cell(i,j).value)
            if len(list_i) >= 1:
                narray = np.array(list_i)
                sheet1.write(x*4+1,j+1,len(list_i))#记录计数
                sheet1.write(x*4+2,j+1,narray.mean())#记录均值
                sheet1.write(x*4+3,j+1,narray.var())#记录方差
        x += 1
        print(x)
    workbook.close()

    t2 = time.time()
    print("%d群组均值融合完成，耗时："%k+str(t2-t1)+"秒。") #反馈结果

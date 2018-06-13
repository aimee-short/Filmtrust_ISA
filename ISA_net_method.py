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
import sys
reload(sys)
sys.setdefaultencoding('utf8')
def ISA_group_ave_net(t,k):
    t1 = time.time()
    #所属群组
    t_group = xlrd.open_workbook('sta\\Filmtrust_isa_%d_%d_ave.xlsx'%(t,k),'rb')#融合后评分 四行 项目dict 计数 均值 方差
    a_group = t_group.sheet_by_name('ave')

    # 每个群组的用户词典
    t_dict = xlrd.open_workbook('sta\\Filmtrust_isa_%d_%d_sta.xlsx'%(t,k),'rb')#三行 第一行：用户数项目数 第二行：用户dict 第三行：项目dict
    group_dict = t_dict.sheet_by_name('sta')

    # 1. 创建一个Excel文件
    workbook = xlsxwriter.Workbook('sta\\Filmtrust_isa_%d_%d_ave_net.xlsx'%(t,k))
    sheet1 = workbook.add_worksheet('ave_net')
    sheet11 = workbook.add_worksheet('ave_net_all')

    # item评分列表
    set_item = set()
    n_item = []#每个群组item数量
    n_user = []
    for i in range(k):
        sheet1.write(i+1,0,'G_%d'%(i+1))#写入抬头 群组ID
        sheet11.write(i+1,0,'G_%d'%(i+1))#写入抬头 群组ID
        n_item.append(int(group_dict.cell(i+1,3).value))#每个群组中item数量
        n_user.append(int(group_dict.cell(i+1,2).value))
    #为啥要费劲巴拉地遍历读取一遍，直接全部item不就好了？这样读取最后预测的时候还要重新对照一遍item   ID  脑子秀逗了吧！！！！傻逼！！！！

    # 写入抬头 项目ID
    for i in range(2701):
        sheet1.write(0,i+1,'I_%d'%(i+1))
        sheet11.write(0,i+1,'I_%d'%(i+1))

    for i in range(k):
        for j in range(2701):
            sheet1.write(i+1,j+1,0)

    # 写入群组评分 构建群组评分网络
    for ii in range(k):
        for jj in range(n_item[ii]):
            #try:
            ads = int(a_group.cell(ii*4,jj+1).value)# 检索群组ii的项目所在总项目的位置 左侧有title 所以+1
            n = a_group.cell(ii*4+1,jj+1).value
            try:#出错原因：有些群组中虽然包含user跟item，但其实并没有评分内容，需要判断
                if n > n_user[ii]*0.3:# 引入条件 如果评分人数少于组内70%，均值不加入群组虚拟用户中
                    sheet1.write(ii+1,ads,float(a_group.cell(ii*4+2,jj+1).value))# 写入excel中有title  ads群组ii+1  ads刚好
                if a_group.cell(ii*4+2,jj+1).ctype == 2:
                    sheet11.write(ii+1,ads,float(a_group.cell(ii*4+2,jj+1).value))
            except:
                print('%d  %d'%(ii+1,ads) + a_group.cell(ii*4+2,jj+1).value + '出错')

        print(ii)
    workbook.close()
    t2 = time.time()
    print("%d群组均值网络构建完成，耗时："%k+str(t2-t1)+"秒。") #反馈结果

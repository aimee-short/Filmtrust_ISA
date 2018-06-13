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
import sys
reload(sys)
sys.setdefaultencoding('utf8')
def ISA_save(t,k):
    t1 = time.time()

    #是否曾经评过分
    #b_rating = xlrd.open_workbook('Filmtrust_array_train_242.xlsx','rb')
    b_rating = xlrd.open_workbook(r'train\\Filmtrust_array_train_%d.xlsx'%t,'rb')
    table = b_rating.sheet_by_index(0)#通过获取
    nrows = table.nrows
    ncols = table.ncols
    name = 'result_%d_%d'%(t,k)

    #matfn = u'result_242_798.mat'#__header__  __version__    result_242_780   __globals
    data = sio.loadmat(r'ISA\\%s.mat'%name)

    # 创建聚类总体信息
    #workbook = xlsxwriter.Workbook('Filmtrust_isa_242_%d_sta.xlsx'%k)
    workbook = xlsxwriter.Workbook('sta\\Filmtrust_isa_%d_%d_sta.xlsx'%(t,k))
    sheet1 = workbook.add_worksheet('array')#第一个表存储评分所属群组
    sheet2 = workbook.add_worksheet('dict')#第二个表存储每个群组的用户跟item信息
    sheet3 = workbook.add_worksheet('sta')#第二个表存储每个群组的用户跟item信息

    # 创建聚类具体信息表
    #workgroup = xlsxwriter.Workbook('Filmtrust_isa_242_%d_all.xlsx'%k)
    workgroup = xlsxwriter.Workbook('sta\\Filmtrust_isa_%d_%d_all.xlsx'%(t,k))

    # 第一张表的群组分类信息
    list_g = []
    n_r = 0
    # 初始化0
    for ii in range(1,nrows):
        for jj in range(1,ncols):
            sheet1.write(ii,jj,0)
    # label
    for i1 in range(1,nrows):
        sheet1.write(i1,0,"U_%d"%i1)
    for j1 in range(1,ncols):
        sheet1.write(0,j1,"I_%d"%j1)
    # 分组信息
    for i in range(k):
        #list_u = data[name]['Clust'][0,0][0,i][0]# 存储行dict
        #list_i = data[name]['Clust'][0,0][0,i][1][0]# 存储列dict
        list_u = data[name]['Clust'][0,0][0,i][0]
        list_i = data[name]['Clust'][0,0][0,i][1]
        sheetg = workgroup.add_worksheet('%d'%(i+1))
        '''
        # 初始化抬头
        for j in range(len(list_u)):
            sheetg.write(j+1,0,list_u[j,0])
        for x in range(len(list_i[0])):
            sheetg.write(0,x+1,list_i[0,x])
        '''
        # 初始化0
        for jj in range(len(list_u)):
            for xx in range(len(list_i[0])):
                sheetg.write(jj,xx,0)
        # 判断并赋值  不要抬头，方便直接读入Matlab
        for j in range(len(list_u)):
            for x in range(len(list_i[0])):
                try:
                    rat = table.cell(list_u[j,0]-1,list_i[0,x]-1).value
                    if rat != 0:
                        sheet1.write(list_u[j,0],list_i[0,x],i+1)
                        sheetg.write(j,x,rat)
                        n_r += 1
                    else:
                        continue
                except:
                    print('%d报错'%i)
                    continue
        list_g.append(n_r)
        n_r = 0
        print(i)

    format1 = workbook.add_format({'bold': True, 'bg_color' : 'yellow'})
    format2 = workbook.add_format({'bg_color' : 'cyan'})
    format_r = workbook.add_format({'bg_color' : '#FF5151'})
    format3 = workbook.add_format({'bg_color' : '#ADADAD'})

    sheet3.write(0,0,'群组ID')
    sheet3.write(0,1,'评分数量')# 评分数量
    sheet3.write(0,2,'用户数量')# 用户数量
    sheet3.write(0,3,'项目数量')# 项目数量
    sheet3.write(0,4,'群组密度')# 项目数量
    # 分组信息
    for i in range(k):
        #list_u = data[name]['Clust'][0,0][0,i][0]# 存储行dict
        #list_i = data[name]['Clust'][0,0][0,i][1][0]# 存储列dict
        list_u = data[name]['Clust'][0,0][0,i][0]
        list_i = data[name]['Clust'][0,0][0,i][1]
        len_u = len(list_u)
        len_i = len(list_i[0])# list_i[0]是真实item列表
        sheet2.write(3*i,0,i+1, format1)
        sheet2.write(3*i,1,list_g[i],format2)# 评分数量
        sheet2.write(3*i,2,len_u,format3)# 用户数量
        sheet2.write(3*i,3,len_i,format3)# 项目数量
        sheet3.write(i+1,0,i+1)
        sheet3.write(i+1,1,list_g[i])# 评分数量
        sheet3.write(i+1,2,len_u)# 用户数量
        sheet3.write(i+1,3,len_i)# 项目数量
        sheet3.write(i+1,4,float(list_g[i])/(float(len_u)*float(len_i)))# 项目数量
        for j in range(len_u):
            sheet2.write(3*i+1,j,list_u[j,0])
        for x in range(len_i):
            sheet2.write(3*i+2,x,list_i[0,x])
        print(i)

    workbook.close()
    workgroup.close()
    t2 = time.time()
    print("%d聚类词典划分完成，耗时："%k+str(t2-t1)+"秒。") #反馈结果

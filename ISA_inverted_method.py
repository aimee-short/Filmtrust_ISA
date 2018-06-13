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
def PearsonS(vec1, vec2):
    value = range(len(vec1))

    sum_vec1 = sum([ vec1[i] for i in value])
    sum_vec2 = sum([ vec2[i] for i in value])

    square_sum_vec1 = sum([ pow(vec1[i],2) for i in  value])
    square_sum_vec2 = sum([ pow(vec2[i],2) for i in  value])

    product = sum([ vec1[i]*vec2[i] for i in value])

    numerator = product - (sum_vec1 * sum_vec2 / len(vec1))
    dominator = ((square_sum_vec1 - pow(sum_vec1, 2) / len(vec1)) * (square_sum_vec2 - pow(sum_vec2, 2) / len(vec2))) ** 0.5

    if dominator == 0:
        return 0
    result = numerator / (dominator * 1.0)

    return abs(result)

def ISA_group_ave_Pearson(t,k,uid,x):
    #t 训练id    k ISA划分群组数  uid  第几个用户   x 第几个群组
    #所属群组
    a_group = xlrd.open_workbook('sta\\Filmtrust_isa_%d_%d_all.xlsx'%(t,k),'rb')
    gr = a_group.sheet_by_index(x-1)# 评分矩阵

    # 每个群组的用户词典
    group_dict = xlrd.open_workbook('sta\\Filmtrust_isa_%d_%d_sta.xlsx'%(t,k),'rb')
    t_dict = group_dict.sheet_by_name('dict')

    # 群组均值
    t_group = xlrd.open_workbook('sta\\Filmtrust_isa_%d_%d_ave.xlsx'%(t,k),'rb')#融合后评分 四行 项目dict 计数 均值 方差
    t_ave = t_group.sheet_by_name('ave')

    # 读取群组x的均值向量
    n_item = int(t_dict.cell((x-1)*3,3).value)#  群组x中的项目数量
    list_ave = []
    for i in range(n_item):
        list_ave.append(t_ave.cell(4*(x-1)+2,i+1).value)

    #查找用户u在第gi群组中的位置
    n_user = int(t_dict.cell((x-1)*3,2).value)#用户个数
    list_u = []
    for i in range(n_user):
        if int(t_dict.cell((x-1)*3+1,i).value) == uid:
            u_add = i
            break

    # 读取用户uid的评分向量 从0开始
    for i in range(n_item):
        list_u.append(gr.cell(u_add,i).value)

    # 计算Pearson相似度
    sim = PearsonS(list_u,list_ave)
    return(sim)

def ISA_group_inverted(t,k):
    t1 = time.time()

    #是否曾经评过分
    b_rating = xlrd.open_workbook(r'train\\Filmtrust_array_train_%d.xlsx'%t,'rb')
    table = b_rating.sheet_by_index(0)#通过获取
    nrows = table.nrows
    ncols = table.ncols

    # 每个群组的用户词典
    group_dict = xlrd.open_workbook(r'sta\\Filmtrust_isa_%d_%d_sta.xlsx'%(t,k),'rb')
    t_dict = group_dict.sheet_by_name('dict')

    # 评分矩阵
    a_group = xlrd.open_workbook('sta\\Filmtrust_isa_%d_%d_all.xlsx'%(t,k),'rb')

    # 群组均值
    t_group = xlrd.open_workbook('sta\\Filmtrust_isa_%d_%d_ave.xlsx'%(t,k),'rb')#融合后评分 四行 项目dict 计数 均值 方差
    t_ave = t_group.sheet_by_name('ave')

    def ave_Pearson(n_user,n_item,uid,x):
        gr = a_group.sheet_by_index(x)# 评分矩阵
        # 读取群组x的均值向量
        list_ave = []
        for i in range(n_item):
            n = t_ave.cell(4*x+1,i+1).value
            if n < n_user*0.7:# 引入条件 如果评分人数少于组内70%，均值不加入群组虚拟用户中
                list_ave.append(t_ave.cell(4*x+2,i+1).value)
            else:
                list_ave.append(0)

        #查找用户u在第gi群组中的位置
        list_u = []
        for i in range(n_item):
            if int(t_dict.cell(x*3+1,i).value) == uid:
                u_add = i
                break

        # 读取用户uid的评分向量 从0开始
        try:
            for i in range(n_item):
                list_u.append(gr.cell(u_add,i).value)
            # 计算Pearson相似度
            sim = PearsonS(list_u,list_ave)
            return(sim)
        except:
            return(0)

    # 创建预测评分矩阵
    workbook = xlsxwriter.Workbook(r'sta\Filmtrust_isa_%d_%d_inverted.xlsx'%(t,k))
    sheet_u = workbook.add_worksheet('user_ID')#用户-群组索引
    sheet_i = workbook.add_worksheet('item_ID')#项目-群组索引
    sheet_uP = workbook.add_worksheet('user_Pearson')#用户-群组Pearson相似度

    # 写入用户-群组倒排档
    list_ug = [set() for i in range(nrows)]#从零开始
    list_ig = [set() for i in range(ncols)]#从零开始
    sheet_u.write(0,0,'用户ID')
    sheet_u.write(0,1,'用户所在群组数量')
    sheet_u.write(0,2,'用户所在群组ID')
    sheet_i.write(0,0,'项目ID')
    sheet_i.write(0,1,'项目所在群组数量')
    sheet_i.write(0,2,'项目所在群组ID')
    sheet_uP.write(0,0,'用户ID')
    sheet_uP.write(0,1,'用户所在群组数量')
    sheet_uP.write(0,2,'用户与群组Pearson')

    for g in range(k):# 倒排档
        nu = int(t_dict.cell(g*3,2).value)# 群组用户数
        ni = int(t_dict.cell(g*3,3).value)# 群组项目数
        for uid in range(nu):
            aim_u = int(t_dict.cell(g*3+1,uid).value)# 目标user id
            o = len(list_ug[aim_u-1])#数组中要-1
            sheet_u.write(aim_u,o+2,g+1)#写入excel中有title 所以下标是对应的  源数据是从零开始的，无title
            ug_Pearson = ave_Pearson(nu,ni,aim_u,g)# 调用函数计算用户与虚拟均值的Pearson相似度
            sheet_uP.write(aim_u,o+2,ug_Pearson)# 写入对应位置
            list_ug[aim_u-1].add(g+1)
        for iid in range(ni):
            aim_i = int(t_dict.cell(g*3+2,iid).value)
            p = len(list_ig[aim_i-1])
            sheet_i.write(aim_i,p+2,g+1)
            list_ig[aim_i-1].add(g+1)
        print(g)

    for i in range(nrows):# 写入user id
        sheet_u.write(i+1,0,i+1)
        sheet_uP.write(i+1,0,i+1)
        sheet_u.write(i+1,1,len(list_ug[i]))
    for i in range(ncols):# 从0开始 项目id写入
        sheet_i.write(i+1,0,i+1)
        sheet_i.write(i+1,1,len(list_ig[i]))

    workbook.close()
    t2 = time.time()
    print("%d倒排档统计完成，耗时："%k+str(t2-t1)+"秒。") #反馈结果

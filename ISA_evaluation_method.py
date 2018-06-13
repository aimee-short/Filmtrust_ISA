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
import math
import sys
reload(sys)
sys.setdefaultencoding('utf8')

def ndcg(y_truth, y_pred):
    result = []
    i = 0
    for y1_i in y_truth:
      y2_i, i = y_pred[i], i+1
      result.append( (y2_i, y1_i) )
    sort_list = sorted(result, key = lambda d:d[0], reverse=True)
    dcg, i = 0, 0
    for y2_i, y1_i in sort_list:
      i += 1
      dcg += (math.pow(2, y1_i) - 1) / math.log(i+1, 2)
    sort_list = sorted(result, key = lambda d:d[1], reverse=True)
    idcg, i = 0, 0
    for y2_i, y1_i in sort_list:
      i += 1
      idcg += (math.pow(2, y1_i) - 1) / math.log(i+1, 2)
    if idcg == 0:
        return 0
    else:
        return dcg / idcg

def ISA_group_evaluation(t,k):
    t1 = time.time()

    # 测试集向量
    work_test = xlrd.open_workbook(r'train\\Filmtrust_test_split_%d.xlsx'%(t),'rd')
    t_tfk = work_test.sheet_by_name('final_top')#最终top-n比对数据  0:User_ID	 1:Count	 2:Sorted
    t_tfkr = work_test.sheet_by_name('final_top_rating')#最终top-n比对数据 评分结果
    n_u = t_tfk.nrows# 测试用户数量 + 1 因为有title

    # 创建预测评分矩阵
    work_fore = xlrd.open_workbook(r'eva\\Filmtrust_isa_%d_%d_ave_forecast.xlsx'%(t,k),'rd')
    t_fk = work_fore.sheet_by_name('top-k')#最终top-n比对数据  0:User_ID  1:Count_forecast  2:Count_effec  3:Sorted_forecast

    workbook = xlsxwriter.Workbook(r'eva\\Filmtrust_isa_%d_%d_ave_nDCG.xlsx'%(t,k))
    sheet1 = workbook.add_worksheet('eva')

    sheet1.write(0,0,'User_ID')
    sheet1.write(0,1,'Count')
    sheet1.write(0,2,'nDCG')
    sum = 0# nDCG总数
    nf = 0# nDCG计数
    for i in range(1,n_u):# 写入user id
        list_f = []
        list_ff = []
        list_t = []# 从训练集中读取的全部排序项目
        list_tt = []# 从训练集中筛选出预测值中的排序项目
        n_f = int(t_fk.cell(i,2).value)
        if  n_f >= 10:
            for j in range(n_f):# 读取所有的排序项目    或者top-10
                list_f.append(int(t_fk.cell(i,j+3).value))
            for jj in range(int(t_tfk.cell(i,1).value)):# 读取训练集中所有的排序项目
                list_t.append(int(t_tfk.cell(i,jj+2).value))
            for ij in list_f:# 如果训练集中项目在预测集中，挑选出来
                if ij in list_t:
                    ind = list_t.index(ij)
                    list_ff.append(int(t_tfkr.cell(i,ind+2).value))
                else:
                    list_ff.append(0)
            list_tt = sorted(list_ff,reverse=True)
            e = ndcg(list_tt,list_ff)
            sheet1.write(nf+1,0,i)
            sheet1.write(nf+1,1,len(list_ff))
            sheet1.write(nf+1,2,e)
            sum += e
            nf += 1

    sheet1.write(nf+1,0,'Count')
    sheet1.write(nf+1,1,'Ave')

    sheet1.write(nf+2,0,nf)
    sheet1.write(nf+2,1,sum/nf)

    workbook.close()
    print(nf)
    print(sum/nf)
    t2 = time.time()
    print("%dnDCG评价完成，耗时："%k+str(t2-t1)+"秒。") #反馈结果

if __name__ == "__main__":
    ISA_group_evaluation(651,402)
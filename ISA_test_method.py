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
import random
import time
import scipy.io as sio
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

def To_Excel(t):
    f = open(r'train\\FilmTrust_ratings_test_%d.txt'%t,'r')
    lines = f.readlines()

    workbook = xlsxwriter.Workbook(r'train\\Filmtrust_test_%d.xlsx'%t)
    sheet1 = workbook.add_worksheet('test')

    i = 0
    for line in lines:
        a = line.split(' ')
        r = a[0]#行
        c = a[1]#列
        v = a[2]#值
        sheet1.write(i,0,int(r))
        sheet1.write(i,1,int(c))
        sheet1.write(i,2,float(v)*2)
        i += 1
        print(r)
    workbook.close()

def Test_split(t):
    t1 = time.time()
    To_Excel(t)

     # 每个群组的用户词典
    group_dict = xlrd.open_workbook('train\\Filmtrust_test_%d.xlsx'%t,'rb')
    t_test = group_dict.sheet_by_name('test')
    nrows = t_test.nrows

    workbook = xlsxwriter.Workbook(r'train\\Filmtrust_test_split_%d.xlsx'%(t))
    t_all = workbook.add_worksheet('all')
    t_old = workbook.add_worksheet('old')#测试所需对比数据
    t_ov = workbook.add_worksheet('old_vector')#测试所需对比数据 向量
    t_final = workbook.add_worksheet('final')#最终top-n比对数据
    t_fk = workbook.add_worksheet('final_top')#最终top-n比对数据
    t_fkr = workbook.add_worksheet('final_top_rating')#最终top-n比对数据 评分结果

    list_u = []
    uid = t_test.cell(0,0).value
    ma = 0#all写入行数记录
    mo = 0#old写入行数记录
    mov = 1#old_vector写入行数记录  还有final_top  两个表对应用户id是一样的 有title  所以要向下平移一位
    mf = 0#final写入行数记录

    t_fk.write(0,0,'User_ID')
    t_fk.write(0,1,'Count')
    t_fk.write(0,2,'Sorted')
    t_fkr.write(0,0,'User_ID')
    t_fkr.write(0,1,'Count')
    t_fkr.write(0,2,'Sorted')
    for i in range(nrows):
        a = t_test.cell(i,0).value
        if a == uid:#id相同
            list_u.append(i)
        else:#id不同，判断user评分记录是否超过20条，超过，写入all，并划分测试集，否则舍弃
            n_p = len(list_u)
            if n_p > 20:
                rp = n_p * 0.4# 测试集列表
                list_test = random.sample([y for y in range(int(n_p))],int(rp))
                t_ov.write(mov,0,uid)#测试向量写入用户id
                for x in range(2701):#向量填充0
                    t_ov.write(mov,x+1,0)
                t_fk.write(mov,0,uid)#最后对比写入用户id
                t_fkr.write(mov,0,uid)#最后对比写入用户id
                dict_uf = {}# 存储最终测试集item跟rating
                for j in range(n_p):
                    t_all.write(ma,0,t_test.cell(list_u[j],0).value)
                    t_all.write(ma,1,t_test.cell(list_u[j],1).value)
                    t_all.write(ma,2,t_test.cell(list_u[j],2).value)
                    ma += 1
                    if j in list_test:#在0.4中，写入old中
                        t_old.write(mo,0,t_test.cell(list_u[j],0).value)
                        t_old.write(mo,1,t_test.cell(list_u[j],1).value)
                        t_old.write(mo,2,t_test.cell(list_u[j],2).value)
                        mo += 1
                        t_ov.write(mov,int(t_test.cell(list_u[j],1).value),int(t_test.cell(list_u[j],2).value))#因为左边写入用户id，所以列刚好
                    else:
                        t_final.write(mf,0,t_test.cell(list_u[j],0).value)
                        t_final.write(mf,1,t_test.cell(list_u[j],1).value)
                        t_final.write(mf,2,t_test.cell(list_u[j],2).value)
                        mf += 1
                        dict_uf[int(t_test.cell(list_u[j],1).value)] = int(t_test.cell(list_u[j],2).value)
                dict_uk = sorted(dict_uf.items(),key = lambda x:x[1],reverse = True)#降序
                t_fk.write(mov,1,len(dict_uf))
                t_fkr.write(mov,1,len(dict_uf))
                fk = 2
                for key,value in dict_uk:# 将排序结果写入后两个表中
                    t_fk.write(mov,fk,key)
                    t_fkr.write(mov,fk,value)
                    fk += 1
                mov += 1
            list_u = []
            uid = t_test.cell(i,0).value
            list_u.append(i)

    # for循环结束之后，需要判断list_u 中最后一个user是否满足条件
    n_p = len(list_u)
    if n_p > 20:
        rp = n_p * 0.4# 测试集列表
        list_test = random.sample([y for y in range(int(n_p))],int(rp))
        t_ov.write(mov,0,uid)
        t_fk.write(mov,0,uid)
        dict_uf = {}# 存储最终测试集item跟rating
        for j in range(n_p):
            t_all.write(ma,0,t_test.cell(list_u[j],0).value)
            t_all.write(ma,1,t_test.cell(list_u[j],1).value)
            t_all.write(ma,2,t_test.cell(list_u[j],2).value)
            ma += 1
            if j in list_test:#在0.4中，写入old中
                t_old.write(mo,0,t_test.cell(list_u[j],0).value)
                t_old.write(mo,1,t_test.cell(list_u[j],1).value)
                t_old.write(mo,2,t_test.cell(list_u[j],2).value)
                mo += 1
                t_ov.write(mov,int(t_test.cell(list_u[j],1).value),int(t_test.cell(list_u[j],2).value))#因为左边写入用户id，所以列刚好
            else:
                t_final.write(mf,0,t_test.cell(list_u[j],0).value)
                t_final.write(mf,1,t_test.cell(list_u[j],1).value)
                t_final.write(mf,2,t_test.cell(list_u[j],2).value)
                mf += 1
                dict_uf[int(t_test.cell(list_u[j],1).value)] = int(t_test.cell(list_u[j],2).value)
        dict_uk = sorted(dict_uf.items(),key = lambda item:item[1],reverse = True)#降序
        fk = 1
        for key,value in dict_uk:# 将排序结果写入后两个表中
            t_fk.write(mov,fk,key)
            t_fkr.write(mov,fk,value)
            fk += 1
        mov += 1
    print(mov)
    workbook.close()
    t2 = time.time()
    print("%d测试集划分完成，耗时："%t+str(t2-t1)+"秒。") #反馈结果

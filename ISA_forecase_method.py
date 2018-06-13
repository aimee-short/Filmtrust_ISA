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
    #测试集
    work_test = xlrd.open_workbook(r'train\\Filmtrust_test_split_65%d.xlsx'%(t),'rb')
    t_ov = work_test.sheet_by_name('old_vector')#测试所需对比数据 向量
    t_fk = work_test.sheet_by_name('final_top')#最终top-n比对数据
    t_fkr = workbook.sheet_by_name('final_top_rating')#最终top-n比对数据 评分结果

    # 群组均值网络
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

def ISA_group_forecast(t,k):
    t1 = time.time()

    #是否曾经评过分
    b_rating = xlrd.open_workbook(r'train\\Filmtrust_array_train_%d.xlsx'%t,'rb')
    table = b_rating.sheet_by_index(0)#通过获取
    nrows = table.nrows
    ncols = table.ncols

    # 每个群组的均值网络
    group_net = xlrd.open_workbook(r'sta\\Filmtrust_isa_%d_%d_ave_net.xlsx'%(t,k),'rb')
    g_net = group_net.sheet_by_name('ave_net')# 两边都有title  I_1   G_1
    g_net_all = group_net.sheet_by_name('ave_net_all')# 两边都有title  I_1   G_1

    # 倒排档
    group_inverted = xlrd.open_workbook(r'sta\Filmtrust_isa_%d_%d_inverted.xlsx'%(t,k),'rb')
    g_inv = group_inverted.sheet_by_name('item_ID')

    # 测试集向量
    work_test = xlrd.open_workbook(r'train\\Filmtrust_test_split_%d.xlsx'%(t),'rd')
    t_ov = work_test.sheet_by_name('old_vector')#测试所需对比数据 向量
    t_fk = work_test.sheet_by_name('final_top')#最终top-n比对数据
    t_fkr = work_test.sheet_by_name('final_top_rating')#最终top-n比对数据 评分结果
    n_u = t_ov.nrows-1# 测试用户数量 因为首行为空，所以使用时应该减一

    # 创建预测评分矩阵
    workbook = xlsxwriter.Workbook(r'eva\\Filmtrust_isa_%d_%d_ave_forecast.xlsx'%(t,k))
    sheet_P = workbook.add_worksheet('Pearson')#用户-群组Pearson相似度
    sheet_f = workbook.add_worksheet('forecast')
    sheet_k = workbook.add_worksheet('top-k')
    sheet_r = workbook.add_worksheet('top-rating')

    for i in range(n_u):# 写入user id
        sheet_P.write(i,0,t_ov.cell(i,0).value)
        sheet_f.write(i,0,t_ov.cell(i,0).value)
        sheet_k.write(i,0,t_ov.cell(i,0).value)
        sheet_r.write(i,0,t_ov.cell(i,0).value)
    for i in range(k):# 从0开始 项目id写入
        sheet_P.write(0,i+1,'G_%d'%(i+1))

    # 分别读取user vector跟 group vector
    list_g = [[] for j in range(ncols)]
    list_u = [[] for j in range(ncols)]
    for i in range(k):
        for j in range(ncols):
            list_g[i].append(float(g_net.cell(i+1,j+1).value))

    for i in range(n_u):
        for j in range(ncols):
            try:
                list_u[i].append(int(t_ov.cell(i+1,j+1).value))#用user id 有item抬头 空了一行
            except:
                break
    # 计算Pearson相似度
    list_P = [[] for i in range(k)]
    for i in range(n_u):
        for j in range(k):
            try:
                p = PearsonS(list_u[i],list_g[j])
            except:
                p = 0
            sheet_P.write(i+1,j+1,p)
            list_P[i].append(p)
    for i in range(n_u):#第几个用户
        dict_u = {}
        for j in range(ncols):# 项目 下标 241 实际 id 242
            if t_ov.cell(i,j+1).value != 0:#如果该项目没有被评过分
                ng = int(g_inv.cell(i+1,1).value)#每个item所属群组数
                sum = 0.0
                if ng > 0:
                    for ii in range(ng):#遍历item所在的每个群组   item + group  未必包含user
                        gid = int(g_inv.cell(i+1,ii+2).value)#群组id
                        if g_net_all.cell(gid,j+2).ctype == 0:#有些群组中虽然包含user跟item，但其实并没有评分内容，需要判断
                            continue
                        else:
                            sum += float(list_P[i][gid-1])*g_net_all.cell(gid,j+2).value#未评分项目相似度 * 群组均值
                    sheet_f.write(i,j+1,sum)
                    dict_u[j] = sum
        sheet_k.write(0,0,'User_ID')
        sheet_k.write(0,1,'Count_forecast')
        sheet_k.write(0,2,'Count_effec')
        sheet_k.write(0,3,'Sorted_forecast')
        sheet_r.write(0,0,'User_ID')
        sheet_r.write(0,1,'Count_forecast')
        sheet_r.write(0,2,'Count_effec')
        sheet_r.write(0,3,'Sorted_rating')
        sheet_k.write(i+1,1,len(dict_u))
        sheet_r.write(i+1,1,len(dict_u))
        dict_k = sorted(dict_u.items(),key = lambda item:item[1],reverse = True)#降序
        fk = 3
        nn = 0
        for key,value in dict_k:# 将排序结果写入后两个表中
            sheet_k.write(i+1,fk,key+1)
            sheet_r.write(i+1,fk,value)
            if value != 0:
                nn += 1
            fk += 1
        sheet_k.write(i+1,2,nn)
        sheet_r.write(i+1,2,nn)
    workbook.close()
    t2 = time.time()
    print("%d预测值计算完成，耗时："%k+str(t2-t1)+"秒。") #反馈结果

if __name__ == "__main__":
    ISA_group_forecast(651,402)

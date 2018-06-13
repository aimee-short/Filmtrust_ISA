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
from ISA_test_method import Test_split# 数据预处理阶段. 训练集进一步划分，在Data_pre 之后的补救
from ISA_method import ISA_save# 1.从ISA调用ISA_save方法，将分组后的mat转存到excel中
from ISA_ave_method import ISA_group_ave# 2. 分组后评分融合，均值
from ISA_net_method import ISA_group_ave_net# 3.群组均值网络构建
from ISA_inverted_method import ISA_group_inverted # 4. 倒排档，计算用户与均值虚拟用户的Pearson相似度
from ISA_forecase_method import ISA_group_forecast# 5. 计数Pearson相似度等
from ISA_evaluation_method import ISA_group_evaluation# 6. Top-k推荐并计算nDCG

import sys
reload(sys)
sys.setdefaultencoding('utf8')

if __name__ == "__main__":

    #list_isa = [780]
    t1 = time.time()
    list_isa_t = [651,652,653,654,655,656,657,658,659,6510]
    list_isa_k = [402,376,417,441,387,388,386,430,400,436]
    # 测试：651,402
    #k1 = [345,349,371,387,388,394,402]
    k1 = [349,371,387,388,394]
    k2 = [376,382,389,398,414,426]
    k3 = [389,408,411,414,417]
    #k4 = [405,406,409,430,441,442]
    k4 = [406,409,430,441,442]
    k5 = [387,405,411,417,428,432]
    k6 = [364,365,370,388,393,399]
    k7 = [317,386,398,408,418,449]
    k8 = [381,402,415,427,430,448]
    k9 = [365,366,400,403,404,409]
    k10 = [354,378,400,419,436,441]

    '''
    # 第一步 群组存储到Excel中
    # 第二步 计算群组评分均值
    # 第三步 构建群组均值网络
    # 第四步 倒排档，计算用户与均值虚拟用户的Pearson相似度
    # 第五步 计数Pearson相似度等
    # 第六步 Top-k推荐并计算nDCG
    ISA_save
    ISA_group_ave
    ISA_group_ave_net
    ISA_group_inverted
    Test_split
    ISA_group_forecast
    ISA_group_evaluation

    Test_split(6510)
    t = 652
    for i in k2:
        ISA_save(t,i)
        ISA_group_ave(t,i)
        ISA_group_ave_net(t,i)
        ISA_group_inverted(t,i)
        ISA_group_forecast(t,i)
        ISA_group_evaluation(t,i)
    t = 653
    for i in k3:
        ISA_save(t,i)
        ISA_group_ave(t,i)
        ISA_group_ave_net(t,i)
        ISA_group_inverted(t,i)
        ISA_group_forecast(t,i)
        ISA_group_evaluation(t,i)

    t = 657
    ISA_group_forecast(t,317)
    ISA_group_evaluation(t,317)

    for i in k4:
        ISA_save(t,i)
        ISA_group_ave(t,i)
        ISA_group_ave_net(t,i)
        ISA_group_inverted(t,i)
        ISA_group_forecast(t,i)
        ISA_group_evaluation(t,i)

    t = 655
    for i in k5:
        ISA_save(t,i)
        ISA_group_ave(t,i)
        ISA_group_ave_net(t,i)
        ISA_group_inverted(t,i)
        ISA_group_forecast(t,i)
        ISA_group_evaluation(t,i)
    t = 656
    for i in k6:
        ISA_save(t,i)
        ISA_group_ave(t,i)
        ISA_group_ave_net(t,i)
        ISA_group_inverted(t,i)
        ISA_group_forecast(t,i)
        ISA_group_evaluation(t,i)
        '''
    t = 657
    for i in k7:
        ISA_save(t,i)
        ISA_group_ave(t,i)
        ISA_group_ave_net(t,i)
        ISA_group_inverted(t,i)
        ISA_group_forecast(t,i)
        ISA_group_evaluation(t,i)
        '''
    t = 658
    for i in k8:
        ISA_save(t,i)
        ISA_group_ave(t,i)
        ISA_group_ave_net(t,i)
        ISA_group_inverted(t,i)
        ISA_group_forecast(t,i)
        ISA_group_evaluation(t,i)
    t = 659
    for i in k9:
        ISA_save(t,i)
        ISA_group_ave(t,i)
        ISA_group_ave_net(t,i)
        ISA_group_inverted(t,i)
        ISA_group_forecast(t,i)
        ISA_group_evaluation(t,i)
    t = 6510
    for i in k10:
        ISA_save(t,i)
        ISA_group_ave(t,i)
        ISA_group_ave_net(t,i)
        ISA_group_inverted(t,i)
        ISA_group_forecast(t,i)
        ISA_group_evaluation(t,i)


    # 第二步 群组评分预测
    for i in list_isa_242:
        ISA_group_ave(t,i)

    # nmf预测评分与测试集直接对比   效果不好，有效值太少
    for j in list_isa_242:
        #ISA_group_ave_net(t,j)
        #NMF_if_suc(t,j)
        ISA_group_nmf_eva_2(t,j)
    '''
    t2 = time.time()
    print("全部聚类词典划分完成，耗时："+str(t2-t1)+"秒。") #反馈结果

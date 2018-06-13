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
import xlrd
import xlwt
import xlsxwriter
import random
import time
import sys
reload(sys)
sys.setdefaultencoding('utf8')
def Data_pre(t):
    t1 = time.time()
    f = open('FilmTrust_ratings.txt','r')
    lines = f.readlines()
    n_r = 1508
    rp = n_r * 0.2# 测试集列表
    list_test = random.sample([y for y in range(int(n_r))],int(rp))

    # 1. 创建一个Excel文件
    workbook = xlsxwriter.Workbook(r'train\\Filmtrust_array_train_65%d.xlsx'%t)
    # 2. 创建一个工作表sheet对象
    sheet1 = workbook.add_worksheet('train')
    f_test = open(r'train\\FilmTrust_ratings_test_65%d.txt'%t,'w')
    # 初始化0
    for i in range(1508):
        for j in range(2071):
            sheet1.write(i,j,0)
    '''
    for kk in range(1509):
        sheet1.write(kk,0,kk)
    for k in range(2072):
        sheet1.write(0,k,k)
    '''
    list_test.sort()#升序
    # 读取每条评分数据
    for line in lines:
        a = line.split(' ')
        r = a[0]#行
        c = a[1]#列
        v = a[2]#值
        if int(r) in list_test:#user id 在测试集中
            f_test.write(line)
        else:
            sheet1.write(int(r)-1,int(c)-1,float(v)*2)
        print(r)
    f_test.close()

    workbook.close()

    t2 = time.time()
    print("矩阵输出完成，耗时："+str(t2-t1)+"秒。") #反馈结果

if __name__ == "__main__":
    for i in range(2,11):
        Data_pre(i)
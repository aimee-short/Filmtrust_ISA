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




if __name__ == "__main__":
    t1 = time.time()
    t = [651,652,653,654,655,656,657,658,659,6510]
    k = [[345,349,371,387,388,394,402],
            [376,382,389,398,414,426],
            [389,408,411,414,417],
            [405,406,409,430,441,442],
            [387,405,411,417,428,432],
            [364,365,370,388,393,399],
            [317,386,398,408,418,449],
            [381,402,415,427,430,448],
            [365,366,400,403,404,409],
            [354,378,400,419,436,441]]

    workbook = xlsxwriter.Workbook(r'eva\\Filmtrust_isa_ave_nDCG_all.xlsx')
    sheet1 = workbook.add_worksheet('all_record')
    sheet2 = workbook.add_worksheet('all_group')

    sheet1.write(0,0,'Train_ID')
    sheet1.write(0,1,'Group_ID')
    sheet1.write(0,2,'User_ID')
    sheet1.write(0,3,'Count')
    sheet1.write(0,4,'nDCG')

    sheet2.write(0,0,'ID')
    sheet2.write(0,1,'Train_ID')
    sheet2.write(0,2,'Group_ID')
    sheet2.write(0,3,'Count')
    sheet2.write(0,4,'nDCG_ave')
    def nDCG_all(t,k,nn,ii):
        # 打开单个评分文件  nn 写入所有记录的第几行    ii 写入第几个群组
        work_fore = xlrd.open_workbook(r'eva\\Filmtrust_isa_%d_%d_ave_nDCG.xlsx'%(t,k),'rd')
        t_nDCG = work_fore.sheet_by_name('eva')#最终top-n比对数据  0:User_ID  1:Count_forecast  2:Count_effec  3:Sorted_forecast
        nrows = t_nDCG.nrows-3# title  最后两行总结
        try:
            for i in range(nrows):
                sheet1.write(nn+i+1,0,t)
                sheet1.write(nn+i+1,1,k)
                sheet1.write(nn+i+1,2,int(t_nDCG.cell(i+1,0).value))
                sheet1.write(nn+i+1,3,int(t_nDCG.cell(i+1,1).value))
                sheet1.write(nn+i+1,4,t_nDCG.cell(i+1,2).value)
        except:
            print t
            print k
            print nn
            print ii
            print i+1

        sheet2.write(ii+1,0,ii)
        sheet2.write(ii+1,1,t)
        sheet2.write(ii+1,2,k)
        sheet2.write(ii+1,3,int(t_nDCG.cell(nrows+2,0).value))
        sheet2.write(ii+1,4,t_nDCG.cell(nrows+2,1).value)
        return(nn+nrows)

    nn = 0
    ii = 1
    for i in range(10):
        for j in range(len(k[i])):
            nn = nDCG_all(t[i],k[i][j],nn,ii)
            ii += 1
    workbook.close()
    t2 = time.time()
    print("全部nDCG统计完成，耗时："+str(t2-t1)+"秒。") #反馈结果

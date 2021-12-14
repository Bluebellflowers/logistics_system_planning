# 提取od信息
import copy
import numpy as np
import pandas as pd
import openpyxl as op
"""
###########储存预测的ai和pi#########
"""
od_data = pd.read_csv('L_OD_data.csv')
od_data = np.array(od_data)
od_data = od_data.tolist()

ai = od_data[-1]  # 当前吸引量
pi = []  # 当前发生量
for i in od_data:
    del i[0]
    pi.append(i[-1])
del ai[-1]
del pi[-1]

area_number = len(od_data[0]) - 1  # 小区数

elastic_coefficient = [[0.12, 0.69], [0.1, 1], [0.08, 0.88]]
alpha = []
for each in elastic_coefficient:
    alpha.append(each[0] * each[1])


def cal_ai_pi(alpha, year):
    ai_pre = copy.deepcopy(ai)
    pi_pre = copy.deepcopy(pi)
    for i in alpha:
        for index, data in enumerate(ai_pre):
            ai_pre[index] = data * pow(1 + i, year)
    for j in alpha:
        for index, data in enumerate(pi_pre):
            pi_pre[index] = data * pow(1 + j, year)
    return ai_pre, pi_pre

excel=[]
stage=0
ai_pre, pi_pre = cal_ai_pi(alpha[0:stage + 1], 5)
excel.append(ai_pre),excel.append(pi_pre)
print(ai_pre,pi_pre)

stage=1
ai_pre, pi_pre = cal_ai_pi(alpha[0:stage + 1], 5)
excel.append(ai_pre),excel.append(pi_pre)
print(ai_pre,pi_pre)

stage=2
ai_pre, pi_pre = cal_ai_pi(alpha[0:stage + 1], 5)
excel.append(ai_pre),excel.append(pi_pre)
print(ai_pre,pi_pre)

wb = op.Workbook()  # 创建工作簿对象
ws = wb['Sheet']  # 创建子表
for i in excel:
    ws.append(i)
wb.save('ai_pi_pre' + '.xlsx')
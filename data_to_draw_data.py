import numpy as np
import pandas as pd
import openpyxl as op
import json  # 解析json数据
import urllib.request  # 发送请求
from urllib import parse  # URL编码
import matplotlib.pyplot as plt

"""
########处理绘制期望线图需要的数据###########
    1.获取两两之间od
"""
# 获取城市名称
with open('city_name.txt', encoding='utf8') as f:
    city_name = f.read()
city_name = city_name.rsplit('\n')
city_name = city_name[0].split('，')


# 获取od矩阵
def get_one_to_one_od(fileName, index):
    data = pd.read_excel(fileName, header=None)
    data = np.array(data)
    data = data.tolist()
    # 获取两两之间od
    od_column = []
    for i in range(len(data)):

        for j in range(len(data)):
            row = []
            row.append(city_name[i]), row.append(city_name[j])
            row.append(data[i][j])
            od_column.append(row)
    wb = op.Workbook()  # 创建工作簿对象
    ws = wb['Sheet']  # 创建子表
    ws.append(['from_node', 'end_node', 'od'])
    for i in od_column:
        ws.append(i)
    fileName1 = 'od_draw' + str(index)
    wb.save(fileName1 + '.xlsx')


# 通过高德api查询城市经纬度
key = "22709a08ff193cc3493ebbdae106c1a3"
city_coordinate = []  # 城市经纬度
for i in city_name:
    parameters = 'key=' + key + '&address=' + i  # 参数
    url = "https://restapi.amap.com/v3/geocode/geo?{}".format(parameters)  # 拼接请求
    newUrl = parse.quote(url, safe="/:=&?#+!$,;'@()*[]")  # 编码
    response = urllib.request.urlopen(newUrl)  # 发送请求
    data = response.read()  # 接收数据
    jsonData = json.loads(data)  # 解析json文件
    name = jsonData["geocodes"][0]["location"]
    city_coordinate.append(name)
# print(city_coordinate)
# 写入
wb = op.Workbook()  # 创建工作簿对象
ws = wb['Sheet']  # 创建子表
ws.append(['name', 'x', 'y'])
for index, i in enumerate(city_coordinate):
    x, y = i.split(',')
    info = [city_name[index], x, y]
    ws.append(info)
fileName = 'city_coordinate1'
wb.save(fileName + '.xlsx')

fileName = 'od_draw_data1.xlsx'
get_one_to_one_od(fileName, 1)

fileName = 'od_draw_data1.xlsx'
get_one_to_one_od(fileName, 2)

fileName = 'od_draw_data1.xlsx'
get_one_to_one_od(fileName, 3)

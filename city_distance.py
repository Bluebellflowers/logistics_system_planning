import json  # 解析json数据
import urllib.request  # 发送请求
from math import ceil, sqrt
from urllib import parse  # URL编码
import openpyxl as op
import copy

"""
###############获取距离矩阵#############
"""
# 获取城市名称
with open('city_name.txt', encoding='utf8') as f:
    city_name = f.read()
city_name = city_name.rsplit('\n')
city_name = city_name[0].split('，')

# 获取城市的面积计算平均直径,作为自己到自己的距离
with open('city_sq.txt', encoding='utf8') as f:
    city_sq = f.read()
city_sq = city_sq.rsplit('\n')
city_sq = city_sq[0].split(',')
pi = 3.141593
city_d = []
for i in city_sq:
    r = sqrt(eval(i) / pi)*1000
    d=2*r*1.3
    city_d.append(d)

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

# 获取任意两城市之间的距离
city_distance = []
for i in city_coordinate:
    temp = copy.deepcopy(city_coordinate)
    temp.remove(i)
    city_part_distance = []
    for j in temp:
        parameters = 'key=' + key + '&origin=' + i + '&destination=' + j  # 参数
        url = "https://restapi.amap.com/v5/direction/driving?{}".format(parameters)  # 拼接请求
        newUrl = parse.quote(url, safe="/:=&?#+!$,;'@()*[]")  # 编码
        response = urllib.request.urlopen(newUrl)  # 发送请求
        data = response.read()  # 接收数据
        jsonData = json.loads(data)  # 解析json文件
        # print(jsonData)
        name = jsonData["route"]['paths'][0]["distance"]
        city_part_distance.append(name)
    city_distance.append(city_part_distance)
    print(city_part_distance)
# print(city_distance)

# 整理数据，插入自己到自己的距离为
for index,data in enumerate(city_distance):
    data.insert(index, city_d[index])

# 将数据写入excel
wb = op.Workbook()  # 创建工作簿对象
ws = wb['Sheet']  # 创建子表
for i in city_distance:
    ws.append(i)
wb.save('city_distance.xlsx')

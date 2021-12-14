import geopandas
import pandas as pd
import matplotlib as mpl
import matplotlib.pyplot as plt
import matplotlib
from shapely.geometry import Point, Polygon, shape

"""
########################绘制期望线图############################
    1.使用data_to_draw_data函数将数据转化为可以用来画图的数据
    2.连接数据
"""


def read_od_data(fileNme):
    # 读取od文件
    data = pd.read_excel(fileNme)
    # 读取坐标文件
    district = pd.read_excel("city_coordinate1.xlsx")
    # 连接两个表格
    district.columns = ['from_node', 'O_x', 'O_y']
    data = pd.merge(data, district, on=['from_node'])
    district.columns = ['end_node', 'D_x', 'D_y']
    data = pd.merge(data, district, on=['end_node'])
    # print(data)
    return data


def draw(data, fileName):
    # 读取底图
    shp = r'C:\Users\12059\PycharmProjects\logistics_system_planning\map\provice.shp'
    hz = geopandas.GeoDataFrame.from_file(shp, encoding='utf-8')
    plt.figure(1, (10, 10), dpi=300)
    ax = plt.subplot(111)
    plt.sca(ax)
    # 绘制行政区划，底图为白色，边框为黑色，宽度为0.5
    hz.plot(ax=ax, edgecolor=(0, 0, 0, 1), facecolor=(0, 0, 0, 0), linewidths=0.5)
    # 设置colormap的数据
    vmax = max(data['od'])
    # 标准化到0-1
    norm = mpl.colors.Normalize(vmin=0, vmax=vmax)
    # 设定colormap的颜色
    cmapname = 'autumn_r'
    cmap = matplotlib.cm.get_cmap(cmapname)
    # 绘制OD
    for i in range(len(data)):
        # 设定第i条线的color和linewidth
        color_i = cmap(norm(data['od'].iloc[i]))
        linewidth_i = norm(data['od'].iloc[i]) * 20
        # 绘制
        plt.plot([data['O_x'].iloc[i], data['D_x'].iloc[i]],
                 [data['O_y'].iloc[i], data['D_y'].iloc[i]],
                 color=color_i, linewidth=linewidth_i)
    plt.imshow([[0, vmax]], cmap=cmap)
    # 设定colorbar的大小和位置
    cax = plt.axes([0.1, 0.2, 0.02, 0.3])
    plt.colorbar(cax=cax)
    # 设置显示区域
    ax.set_xlim(101, 120)
    ax.set_ylim(22, 41)
    plt.axis('off')
    plt.savefig(fileName, dpi=300)
    plt.close()


if __name__ == '__main__':
    fileName1 = 'od_draw1.xlsx'
    draw(read_od_data(fileName1), 'picture1.png')

    fileName2 = 'od_draw2.xlsx'
    draw(read_od_data(fileName2), 'picture2.png')

    fileName3 = 'od_draw3.xlsx'
    draw(read_od_data(fileName3), 'picture3.png')

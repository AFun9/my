# coding: utf-8
# @Author:Afun
import os
import numpy as np
import pandas as pd
import xlrd as xd
from matplotlib import pyplot as plt
plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False
header = {
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/103.0.5060.53 Safari/537.36 Edg/103.0.1264.37"}


def get_list1(path):  # 读取爬取到的数据
    """
    :param path: excel表的路径
    :return: 返回一个列表，包含excel表中的所有信息
    """
    data = xd.open_workbook(path)  # 打开excel表所在路径
    sheet = data.sheet_by_name('page_1')  # 读取数据，以excel表名来打开
    d = []
    l = []
    for r in range(sheet.nrows):  # 将表中数据按行逐步添加到列表中，最后转换为list结构
        data1 = []
        for c in range(sheet.ncols):
            data1.append(sheet.cell_value(r, c))
        d.append(list(data1))
    for i in range(2, len(d)):
        l.append(d[i])
    return l


def findc(l):
    print(l)
    d = {}
    l = np.array(l)
    a = l[:, 3]
    c = list(a)
    for i in a:
        d[i] = c.count(i)
    print(d)
    return d


def draw_from_dict(dicdata, RANGE, heng=0):
    by_value = sorted(dicdata.items(), key=lambda item: item[1], reverse=True)
    x = []
    y = []
    for d in by_value:
        x.append(d[0])
        y.append(d[1])
    if heng == 0:
        plt.bar(x[0:RANGE], y[0:RANGE])
        plt.show()
        return
    elif heng == 1:
        plt.barh(x[0:RANGE], y[0:RANGE])
        plt.show()
        return
    else:
        return "heng的值仅为0或1！"


def fx(grade, a):
    l = get_list1()
    dg = grade - a
    dgmax = dg + 10
    dgmin = dg - 10
    g = []
    for i in l:
        try:
            if dgmin <= (int(i[8]) - int(i[10])) <= dgmax or dgmin <= (int(i[11]) - int(i[13])) <= dgmax or dgmin <= (
                    int(i[14]) - int(i[16])) <= dgmax:
                g.append(i)
        except:
            pass
    return g


def get211or985():
    l = get_list1()
    c = []
    m = []
    d = {}
    for i in l:
        if i[4] == "双一流":
            c.append(i)
    for i in range(len(c)):
        m.append(c[i][3])
    for i in m:
        d[i] = m.count(i)
    return d


def draw985or211(dicdata, RANGE, heng=0):
    by_value = sorted(dicdata.items(), key=lambda item: item[1], reverse=True)
    x = []
    y = []
    for d in by_value:
        x.append(d[0])
        y.append(d[1])
    if heng == 0:
        plt.bar(x[0:RANGE], y[0:RANGE])
        plt.show()
        return
    elif heng == 1:
        plt.barh(x[0:RANGE], y[0:RANGE])
        plt.show()
        return
    else:
        return "heng的值仅为0或1！"


def writerin(l, path):  # 将数据写入excel文件中
    col = [
        ['编号', '高校名称', '所在地', '所在省', '等级', '类型', '性质', '批次', '2021最低录取线', '2021最低录取排名',
         '2021投档线', '2020最低录取线',
         '2020最低排名',
         '2020投档线',
         '2019最低录取线', '2019最低排名', '2019投档线']]
    list = np.array(col + l)
    data = pd.DataFrame(list)
    writer = pd.ExcelWriter(path)
    data.to_excel(writer, 'page_1', float_format='%.5f')
    writer.save()
    writer.close()


path = 'data/data.xlsx'
if not os.path.exists(path):
    os.mkdir(path)
l = get_list1(path)
b = findc(l)
writerin(fx(459, 440))

# coding: utf-8
# @Author:Afun
import json
import os.path
import time
import numpy as np
import pandas as pd
import requests

# 头信息，根据自己浏览器的头进行更改
header = {
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/103.0.5060.53 Safari/537.36 Edg/103.0.1264.37"}


def get_id_infor(page):
    """
    :param page: int型的数，表示需要爬取的页数，可以根据自己的需求填写
    :return: id:返回一个列表，存储了从第一页到当前页的所有高校的id,infor:一个列表，包含各id高校的基本信息
    """
    url = ["https://api.eol.cn/web/api/?admissions=&central=&department=&dual_class=&f211=&f985=&is_doublehigh" \
           "=&is_dual_class=&keyword=&nature=36000&page={}&province_id=&ranktype=&request_type=1&school_type=&size=20" \
           "&type=&uri=apidata/api/gk/school/lists&signsafe=e01ce712c2a5a787bf33dd75a69cbd2d ".format(i) for i in
           range(page)]
    id = []
    infor = []
    for uu in url:
        time.sleep(0.5)
        request = requests.get(uu, headers=header).text
        j = json.loads(request)
        user = j['data']['item']
        for i in user:
            id.append(i['school_id'])
            infor.append([i['name'], i['city_name'], i['province_name'],
                          i['type_name'] + i['nature_name'] + i['dual_class_name'], ])  # 名字，城市名，所在省，学校类型
    return id, infor


def get_grader(id, year):  # 获取指定年份的所有列表中的高校的录取线
    """
    :param id: 一个列表，包含学校的id
    :param year: int型的年份，表示要爬取的年份
    :return: 返回一个列表，包含各个学校的录取线
    """
    urls = [r"https://static-data.gaokao.cn/www/2.0/schoolprovinceindex/{}/{}/36/1/1.json".format(year, i) for i in
            id]
    print(urls)
    grader = []
    for ur in urls:
        try:
            request = requests.get(ur, headers=header)
            test = json.loads(request.text.encode('utf-8'))
            user = test["data"]["item"][0]
            grader.append(
                [user['local_batch_name'], user['min'], user['min_section'],
                 user['proscore']])
        except:
            grader.append([["未在江西招生"], ["未在江西招生"], ["未在江西招生"], ["未在江西招生"]])
    return grader


def get_list(infor, grader_2023, grader_2022, grader_2021, grader_2020, grader_2019):  # 将数据整合成一个列表
    """
    :param infor: 基本信息
    :param grader_2023: 23年的录取线
    :param grader_2022: 22年的录取线
    :param grader_2021: 21年的录取线
    :param grader_2020: 20年的录取线
    :param grader_2019: 19年的录取线
    :return: 返回一个列表，包含各学校近五年的基本情况
    """
    l = []
    for i in range(len(infor)):
        l.append(
            infor[i][:] + grader_2023[i][:] + grader_2022[i][:] + grader_2021[i][:] + grader_2020[i][1:] + grader_2019[
                                                                                                               i][1:])
    return l


def writer_infile(l, path):  # 将数据写入excel文件中
    """
    :param l: 一个列表，包含各个学校的相关信息
    :param path: 文件的保存路径
    """
    col = [
        ['高校名称', '所在地', "所在省", "学校类型", '批次', '2023最低录取线', '2023最低录取排名', '2023投档线',
         '2022最低录取线', '2022最低录取排名', '2022投档线', '2021最低录取线', '2021最低录取排名', '2021投档线',
         '2020最低录取线', '2020最低排名',
         '2020投档线',
         '2019最低录取线', '2019最低排名', '2019投档线']]
    list = np.array(col + l)
    data = pd.DataFrame(list)
    writer = pd.ExcelWriter(path)
    data.to_excel(writer, 'page_1', float_format='%.5f')
    writer.save()
    writer.close()


if __name__ == '__main__':
    path = 'data/data.xlsx'
    if not os.path.exists(path):
        os.mkdir(path)
    else:
        id, infor = get_id_infor()
        grade_2019 = get_grader(id, 2019)
        grade_2020 = get_grader(id, 2020)
        grade_2021 = get_grader(id, 2021)
        grade_2022 = get_grader(id, 2022)
        grade_2023 = get_grader(id, 2023)
        list = get_list(infor, grade_2023, grade_2022, grade_2021, grade_2020, grade_2019)
        print(list)
        writer_infile(list, path)

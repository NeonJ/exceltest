# -*- encoding: utf-8 -*-
import requests  # 导入requests包
from bs4 import BeautifulSoup
import re
import time
from multiprocessing import Pool


# 追加写入
def write_file(path, name):
    with open(path, 'a') as f:
        f.write(name + '\n')


# 初始化文件
def create_file(path):
    with open(path, 'w') as f:
        f.write('')


# 读取文件
def read_file(path):
    file = open(path)
    data = file.read().splitlines()
    return data


# 正则匹配
def text_find(rule, text):
    # 定义正则匹配的规则
    pattern = re.compile(rule)
    # 进行匹配
    result = re.findall(pattern, text)
    return result


# 获得url的text
def get_page_text(url):
    request_page = requests.get(url)
    page_text = request_page.text
    return page_text


def dict_set(key_list, value_list):
    dict = {}
    for x in range(0, len(key_list)):
        dict[key_list[x]] = value_list[x]
    return dict


def crawl(url):
    strhtml = requests.get(url)
    soup = BeautifulSoup(strhtml.text, "html.parser")

    # data = soup.select('#active-batches-table > tbody > tr > td:nth-child(3)')
    data = soup.select('#active-batches-table > tbody > tr > td:nth-child(3)')
    print(data)
    if len(data) > 0:
        for x in data:
            # active-batches-table > tbody > tr:nth-child(7) > td:nth-child(3)
            list = text_find(r"[0-9]+", str(x))
            if len(list) > 1:
                schedule_time = int(list[0])
                if schedule_time > 3600000:  # 3600000ms = 1hour 
                    write_file('job_schedule_check_result.txt', url)


if __name__ == '__main__':
    create_file('job_schedule_check_result.txt')
    check_list = read_file('cas_job_list.txt')
    # 获取main_page
    job_main_page = get_page_text("http://spark.cas.ivy")
    # print(job_main_page)
    # 将页面内容正则匹配的结果放入list
    job_url_list = []
    job_name_list = []
    # print(job_url_list)
    # for x in range(0, len(job_url_list)):
    #     job_url_list[x] = job_url_list[x].replace("\"", "") + '/streaming/'
    # job_name_list = text_find(r"<a href=(.*?)>(.*?)</a>", job_main_page)
    # print(job_name_list)

    soup = BeautifulSoup(job_main_page, "html.parser")
    for link in soup.find_all('a'):
        if link.string in check_list:
            job_url_list.append(link.get('href'))
    for address in range(0, len(job_url_list)):
        job_url_list[address] = job_url_list[address].replace("\"", "") + '/streaming/'

    print('----------')

    for jobname in soup.find_all('a'):
        if jobname.string in check_list:
            job_name_list.append(jobname.string)

    print('----------')
    # 将url_list和name_list放入dict字典
    # dict = dict_set(job_url_list, job_name_list)
    # print(dict)
    dict = dict_set(job_url_list, job_name_list)
    print(dict)
    # pool = Pool()
    # pool.map(crawl, job_url_list)
    for url in job_url_list:
        print(url)
        crawl(url)
    L = read_file('job_schedule_check_result.txt')
    for b in L:
        print(dict[b.replace('\n', '')])
    if len(L) == 0:
        print('no job schedule delay!')
import numpy as np
import re
import requests
from lxml import etree
import os
import time
import xlsxwriter as xw
import json
from bs4 import BeautifulSoup
import csv
headers = {
    'User-Agent':'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/112.0.0.0 Mobile Safari/537.36 Edg/112.0.1722.39',
    'Cookie':'Hm_lvt_0dae59e1f85da1153b28fb5a2671647f=1681710636; Hm_lvt_bfa037370fbbd327cd70336871aea386=1681710659; Hm_lpvt_0dae59e1f85da1153b28fb5a2671647f=1681784876; auth=5D619F3556F7B3F02DE744E2E710521DDE9581C6F6F3A561B7E27E1B512BA1DE1D819B8224EC1F528D2D13D34DFC68329FA17F596A6F48C85072C88BE2717423E72A40849E75107B59A45FDC547B8B88667B25E4CF8421C0D622A9AC50B3FEB38AD403A45846F842151C9A2BE468882082B62C75F29C2C50B29AF94F0DD59B42D3E3411A2F0C443136C727F34B137B1D62271D42A34679C394966F382FBC9D65D796EDC5C93CA5FC42D098D64F3CD82B29977A2DA9F88FB8ACAEFE3A579A115C345C5A1A532242C1DE4B3BECCF3FE4A8D7E8A19C76F624D706098C07C762B65EEA487F34CC86C1796441DFD68D5DB6B27E29E427DDCFAC3957DADEFCDC7CF34EAA4ACFC4C93731BC4418471C29FA788AABB690856FC646FD81757F96C5ED91D9F6A169E8B1C817183148BF14EB380183C327EB4CBC0454333A90F4501E08805C79705DC53455DE37F42AF5DC37D073D679AEF89162DFBD9F715B75BA0CC527FC6489F1688CD90CDA278224C3E06E5E6E92C04CFA441E1FFEF01D94964DFA838EB1E1B9B51E7C98E05F4AE83B39F1B006BF902A8E124D8D4DB0FC264A7E07BA837C7BA78F4D03867E40876493D7E230279CA07E6DAF50EDCAA960A1F5889B196F27E2C864D67776723FB6286B3CFD5E70; Hm_lpvt_bfa037370fbbd327cd70336871aea386=1681821286'
}
url = 'http://advanced.fenqubiao.com/Meso/Index?year=2022'
base_url1 = 'http://advanced.fenqubiao.com/Meso/PageData'
base_url2 = 'http://advanced.fenqubiao.com/Meso/GetJson'
# #base_url ='http://advanced.fenqubiao.com/Macro/Journal?'
responseold = requests.get(url=url, headers=headers)
page_textold = responseold.text

def get_page_text1(url, headers, name,year,num):
    id = ''
    class1 = ''
    title = ""
    issn = ""
    review = False
    eidtions = ''
    data = {
       'draw': str(num),
       'columns[0][data]':'',
       'columns[0][name]':'',
       'columns[0][searchable]': True,
       'columns[0][orderable]': False,
       'columns[0][search][value]':'',
       'columns[0][search][regex]': False,
       'columns[1][data]': title,
       'columns[1][name]': title,
       'columns[1][searchable]': True,
       'columns[1][orderable]': False,
       'columns[1][search][value]':'',
       'columns[1][search][regex]': False,
       'columns[2][data]': issn,
       'columns[2][name]': issn,
       'columns[2][searchable]': True,
       'columns[2][orderable]': False,
       'columns[2][search][value]':'',
       'columns[2][search][regex]': False,
       'columns[3][data]': class1,
       'columns[3][name]': class1,
       'columns[3][searchable]': True,
       'columns[3][orderable]': False,
       'columns[3][search][value]':'',
       'columns[3][search][regex]': False,
       'order[0][column]': '0',
       'order[0][dir]': 'asc',
       'start': str((num-1)*20),
       'length': '20',
       'search[value]':'',
       'search[regex]': False,
       'name': name,
       'year': year

    }
    response = requests.post(url=base_url1, headers=headers, data=data)
    page_text = response.text
    return page_text
def get_page_text2(url, headers, name,year,num):

    data = {
        'name': name,
        'keyword':'',
        'start': str(1+(num-1)*20),
        'length': '20',
        'year': year
        # 'name': name,
        # 'keyword': '',
        # 'strat': str((num-1)*20+1),
        # 'length': '20',
        # 'year': year

    }
    response = requests.post(url=url, headers=headers, data=data)
    page_text = response.text
    return page_text
# def get_abstract(url):
#     response = requests.get(url=url, headers=headers)
#     page_text = response.text
#     tree = etree.HTML(page_text)
#     abstract = tree.xpath('//div[@class="xx_font"]//text()')
#
#     return abstract
#
#
def list_to_str(my_list):
    my_str = "".join(my_list)
    return my_str


def parse_page_text(page_text):
    soup = BeautifulSoup(page_text, "html.parser")
    pattern = re.compile(r"NameCn", re.MULTILINE | re.DOTALL)
    script = soup.find('script', text=pattern)
    list_code = []
    list_name = []
    for match in re.finditer('"Code":"(.*?)"',script.text):
        code = match.group().split(":")[1]
        code = code.replace("\"", "").replace("\"", "")
        list_code.append(code)
    for match in re.finditer('"Name":"(.*?)"',script.text):
        name = match.group().split(":")[1]
        name = name.replace("\"", "").replace("\"", "")
        list_name.append(name)

    tree = etree.HTML(page_text)
    itemyear_lists = tree.xpath('//*[@id="year"]/option')
    year_info = []
    for itemyear_list in itemyear_lists:
        year = list_to_str(itemyear_list.xpath('./text()'))
        year_info.append(year)
    print(year_info)
    print(list_code)

    return year_info,list_code,list_name
#
#
def write_to_excel(workbook, info):

    wb = workbook
    worksheet1 = wb.add_worksheet()  # 创建子表
    worksheet1.activate()  # 激活表

    title = ['序号', '刊名', 'ISSN',
             '分区']  # 设置表头
    worksheet1.write_row('A1', title)  # 从A1单元格开始写入表头

    i = 2  # 从第二行开始写入数据
    for j in range(len(info)):
        insert_data = info[j]
        start_pos = 'A' + str(i)
        # print(insert_data)
        worksheet1.write_row(start_pos, insert_data)
        i += 1
    return True


if __name__ == '__main__':
    # 1、创建一个文件夹
     if not os.path.exists('./paper_info_升级版_小类'):
        os.mkdir('./paper_info_升级版_小类')

    #print(page_textold)
     year_info,code_info,name_info = parse_page_text(page_text=page_textold)
     print(year_info)
     infos =[]
     n = -1
     for name in code_info:
         n += 1
         for yearo in year_info:
             file_name = './paper_info_升级版_小类/文献爬取'
             file_name = file_name + yearo + name_info[n] +'.xlsx'
             print(file_name)
             workbook = xw.Workbook(filename=file_name)
             i = 0
             infoall = []
             year = yearo.split("年")[0]
             page_text1 = get_page_text1(base_url1, headers, name, year,i+1)
             text1 = json.loads(page_text1).get('recordsTotal')
             #page_text2 = get_page_text2(base_url2, headers, name, year)
             print((text1//20)+1)
             for page_num in range(1, (text1//20)+2):
             #for page_num in range(1, 2):
                 infos = []
                 page_text1 = get_page_text1(base_url1, headers, name, year,page_num)
                 #print(page_text1)
                 page_text2 = get_page_text2(base_url2, headers, name, year,page_num)
                 #print(page_text2)
                 text1 = json.loads(page_text1).get('data')
                 text2 = json.loads(page_text2)
                 #for j in range(0,20):
                 length = len(text1)
                 classs = np.zeros(length)
                 for j in range(0,length):
                     classs[j] = int(text2[j]['Class'])
                 classs = np.sort(classs)
                 for j in range(0,length):
                     info = []
                     i += 1
                     num = str(i)
                     title = text1[j]['title']
                     issn = text1[j]['issn']
                     class1 = str(int(classs[j]))
                     info = [z.strip() for z in [num, title, issn, class1]]
                     infos.append(info)
                     j += 1
                 infoall += infos
                     #time.sleep(5)
             write_to_excel(workbook, infoall)
             workbook.close()


    #             # 用+合并成一个列表，不是嵌套列表；用append，会形成嵌套列表
    #             infos += page_info
    #             time.sleep(5)
    #
    #     # 5、按照搜索词，依次写入工作簿
    #          write_to_excel(workbook, infos, name,year)
    # # 6、关闭工作簿
    # workbook.close()
    #
    # print('爬取完成!')

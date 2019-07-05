from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException,NoSuchElementException
import sys,os,random,re,time
import xlwt
import requests
import json
import threading
import queue
from lxml import etree
from requests.exceptions import Timeout
from json.decoder import JSONDecodeError

class GuoJiaXingYong(object):
    def __init__(self):
        start_url = "http://www.gsxt.gov.cn/corp-query-entprise-info-xxgg-100000.html"
        self.driver = webdriver.Chrome()
        self.driver.get(start_url)
        self.queue = queue.Queue()
        self.url = 'http://www.gsxt.gov.cn/affiche-query-area-info-paperall.html?'
        self.cookies = {}
        self.headers = {
            "Accept": "application/json, text/javascript, */*; q=0.01",
            "Accept-Encoding": "gzip, deflate",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
            "Connection": "keep-alive",
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "Host": "www.gsxt.gov.cn",
            "Origin": "http://www.gsxt.gov.cn",
            "User-Agent": "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.131 Safari/537.36",
            "X-Requested-With": "XMLHttpRequest"
        }

    def obtain_info(self):
        # 事不过三,判断左边的名单类型是否加载出来
        for i in range(3):
            try:
                WebDriverWait(self.driver,5).until(EC.element_to_be_clickable((By.CLASS_NAME,'page_left')))
            except TimeoutException:
                self.driver.refresh()
                continue
            break

        # 然后获取左边栏目的class-value值和城市省份的id值，作为post请求使用
        category = self.driver.find_elements_by_xpath('//div[@class="page_left"]//div[not(contains(@child-value,"3"))]')
        category_dict = {}
        for i in category:
            if i.get_attribute('class-value') != None:
                category_dict[i.text] = i.get_attribute('class-value')
        province = self.driver.find_elements_by_xpath('//div[@class="label-list"]/div')
        property_dict = {}
        for i in province:
            name = i.find_element_by_xpath('./label').text
            property_dict[name] = i.get_attribute('id')

        # 获取cookie
        cookie = self.driver.get_cookies()
        for i in cookie:
            self.cookies[i['name']] = i['value']
        return (category_dict,property_dict)
        
    def construct_info(self,c_dict,p_dict):
        # 遍历所有的栏目和类别
        for key,noticeType in c_dict.items():
            # 创建相应的存放信息的文件夹
            if not os.path.exists(key):
                os.mkdir(key)
            # 遍历改栏目所有的城市
            for name,regOrg in p_dict.items():
                # 构建请求的url
                url = self.url + 'noticeType=%s&areaid=100000&noticeTitle=&regOrg=%s' %(noticeType,regOrg)
                self.request_post(key,name,url)
            # 刷新更新cookies
            time.sleep(20)
            self.driver.refresh()
            cookie = self.driver.get_cookies()
            for i in cookie:
                self.cookies[i['name']] = i['value']
            
    def request_post(self,key,name,url):
        print('开始爬取' + key + '_' + name + '的信息' )
        info_list = []
        data = {
                "draw": 1,
                "start": 0,
                "length": 10,
                }
        while True:
            try:
                req = requests.post(url,data=data,headers=self.headers,cookies=self.cookies,timeout=30)
            except Timeout:
                print('请求超时啦，重新请求')
                continue
            try:
                infos = json.loads(req.text)["data"]
            except (KeyError,JSONDecodeError):
                print('没有得到我想要的，准备重新再来一次')
                continue
            if infos == []:
                self.write_excel(info_list,key,name)
                break
            # 如何对都一个地方停止请求
            for i in infos:
                info_list.append(i["noticeTitle"])
            print("第" + str(data['draw']) + '条post请求成功')
            data['draw'] += 1
            data['start'] = 10 * data['draw']-1

    def write_excel(self,info_list,key,sheng):
        namelist = []
        for name in info_list:
            if '将' in name:
                name = name[3:]
            else:
                name = name[2:]
            if '列入' in name:
                index = name.index('列')
                name = name[:index]
            elif '移出' in name:
                index = name.index('移')
                name = name[:index]
            namelist.append(name)
        book = xlwt.Workbook()
        table = book.add_sheet('my_book')
        table.write(0,0,'公司名称')
        for i in range(1,len(namelist)+1):
            table.write(i,0,namelist[i-1])
        book.save(key + '/' + sheng + '.xls')

    def main(self):
        c_dict,p_dict = self.obtain_info()
        self.construct_info(c_dict,p_dict)
        self.driver.quit()
        
if __name__ == "__main__":
    xy = GuoJiaXingYong()
    xy.main()


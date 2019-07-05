# 此文档爬取照片
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
from requests.exceptions import ChunkedEncodingError
from selenium.common.exceptions import NoSuchElementException
import threading 
import queue
import time,xlrd
import sys,os
import requests
from requests.exceptions import ChunkedEncodingError

class SB_photo(object):
    def __init__(self,url):
        self.status = 0
        self.start_page = url.split('=')[1]
        self.queue = queue.Queue()
        self.driver = webdriver.Chrome()
        self.driver.get(url)
        # 如果爬取失败请更换以下的headers,
        self.headers = {
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3",
            "Accept-Encoding": "gzip, deflate",
            "Accept-Language": "zh-CN,zh;q=0.9",
            "Cache-Control": "max-age=0",
            "Connection": "keep-alive",
            "Cookie": "__jsluid=f4d3937e85101b1b4d272eb1f6cb8278",
            "Host": "sbggwj.saic.gov.cn:8000",
            "Upgrade-Insecure-Requests": "1",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.90 Safari/537.36",
            }
    @classmethod
    def read_excel(cls):
        # 读取要爬取的链接
        filelist = next(os.walk('./'))[2]
        for i in filelist:
            if '.xls' in i or '.xlsx' in i:
                book = xlrd.open_workbook(i)
                table = book.sheets()[0]
                link = table.row_values(0)[0]
                return link
    
    # 获取爬取的起始值，读取photo文件
    def read_photo(self):
        if os.path.exists('./' + self.start_page + '期商标图片'):
            filenames = next(os.walk('./' + self.start_page + '期商标图片'))[2]
            num_list = []
            for i in filenames:
                num_list.append(int(i.split('.')[0]))
            return max(num_list)
        else:
            return 0

    def clicks(self,page):
        # 点击查看到图片页面
        try:
             WebDriverWait(self.driver,20).until(EC.visibility_of_all_elements_located((By.XPATH,'//tbody/tr/td/a')))
        except TimeoutException:
            self.driver.quit()
            sys.exit('程序被迫退出,请重新爬取')
        self.driver.find_element_by_xpath('//tr[@class="evenBj"][1]/td[8]/a').click()
        time.sleep(10)
        while True:
            try:
                inputs = self.driver.find_element_by_xpath('//input[@id="nowPage"]')
            except NoSuchElementException:
                time.sleep(1)
                continue
            break
        inputs.clear()
        self.driver.find_element_by_xpath('//input[@id="nowPage"]').send_keys(page+1)
        self.driver.find_element_by_xpath('//input[@id="nowPage"]').send_keys(Keys.ENTER)
        time.sleep(2)

    def create_threading(self,page):
        # 获取总页数
        sum_page = input('请输入总页数的值:')
        while True:
            try:
                sum_page = int(sum_page)
            except ValueError:
                sum_page = input('输入的值有误请重新输入:')
                continue
            break

        # 创建爬取图片的线程
        print('开始爬取照片,请稍等...')
        t = threading.Thread(target=self.request,args=(self.start_page,))
        t.start()
        links = []
        links_d = []
        i = page
        while True:
            link = self.driver.find_element_by_xpath('//li[@class="newclass"]/img').get_attribute('src')
            links_d.append(link)
            # 对下载图片的地址进行去重
            if link not in links:
                links.append(link)
                i += 1
                self.queue.put(link + ',' + str(i))
            # 解决页面的阻塞问题
            if len(links_d) == 5:
                if len(set(links_d)) == 1 or self.queue.qsize() >= 200:
                    links.clear()
                    links_d.clear()
                    self.queue.join()
                    self.driver.refresh()
                    self.clicks(i-1)
                    t = threading.Thread(target=self.request,args=(self.start_page,))
                    t.start()
                else:
                    links_d.clear()
            if i == sum_page or self.status == 1:
                self.driver.quit()
                # 获取当前的时间，查看服务器的关闭时间
                end_time = time.strftime('%m-%d-%H-%M-%S',time.localtime(time.time()))
                print(end_time)
                break
            self.driver.find_element_by_xpath('//div[@class="base_right"]/span').click()
            time.sleep(0.5)
        
    def request(self,page):
        time.sleep(10)
        while True:
            if self.queue.empty():
                break
            url_num = self.queue.get()
            url = url_num.split(',')[0]
            num = url_num.split(',')[1]
            try:
                req = requests.get(url,headers=self.headers)
            except ChunkedEncodingError:
                print('无法请求到数据,爬取被迫停止')
                self.queue.task_done()
                self.status = 1
                break
            html = req.content
            if os.path.exists('./' + page + '期商标图片'):
                pass
            else:
                os.mkdir('./' + page + '期商标图片')
            with open('./' + page + '期商标图片/' + num + '.jpg','wb') as f:
                f.write(html)
            print('第' + num + '照片保存成功')
            time.sleep(0.5)
            self.queue.task_done()

    def main(self):
        page = self.read_photo()
        print('已爬取图片: ' + str(page))
        self.clicks(page)
        self.create_threading(page)
        print('图片爬取完毕')

if __name__ == "__main__":
    link = SB_photo.read_excel()
    sb = SB_photo(link)
    sb.main()


class Mapping:
    def __init__(self, iterable):
        self.items_list = []
        self.__update(iterable)

    def update(self, iterable):
        for item in iterable:
            self.items_list.append(item)

    __update = update

class MappingSubclass(Mapping):

    def update(self, keys, values):
        for item in zip(keys, values):
            self.items_list.append(item)

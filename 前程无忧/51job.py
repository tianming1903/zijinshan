from selenium import webdriver
import time
import xlwt
import requests
from lxml import etree
from requests.exceptions import ConnectionError,Timeout
import threading
import queue
import os

class Job(object):
    headers = {'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3',
        'Accept-Encoding': "gzip, deflate, br",
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
        'Connection': 'keep-alive',
        'Host': 'search.51job.com',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.131 Safari/537.36'
        }

    def __init__(self):
        self.driver = webdriver.Chrome()
        self.driver.get('https://51job.com/')
        self.queue = queue.Queue()
        self.lock = threading.Lock()
        self.th_size = []
        self.namelist = []

    def search(self):
        # 选择条件并获取该条件的链接
        input('是否确定了选择条件,按任意键确认:')
        # 获取地址和行业，作为文件名
        addres = self.driver.find_element_by_xpath('//div[contains(@class,"pointer")][1]/input[1]').get_attribute('value')
        hangye = self.driver.find_element_by_xpath('//div[contains(@class,"pointer")][2]/input[1]').get_attribute('value').replace('/','')
        # 获取链接
        url = self.driver.current_url
        # 获取有多少页
        page = self.driver.find_element_by_xpath('//span[@class="td"][1]').text
        yeshu = ''
        for i in page:
            try:
                int(i)
            except ValueError:
                pass
            else:
                yeshu += i
        page = int(yeshu)
        # 获取cookie
        cookies = {}
        cookie = self.driver.get_cookies()
        for item in cookie:
            cookies[item['name']] = item['value']
        return (addres,hangye,cookies,page,url)

    # 把链接存放到队列里面去
    def put_queue(self,page,url):
         for index in range(1,page+1):
            re_url = url.replace('1.html',str(index) + '.html')
            self.queue.put(re_url)

    def request(self,cookie,headers,url):
        while True:
            try:
                req = requests.get(url,cookies=cookie,headers=headers,timeout=5)
            except (ConnectionError,Timeout):
                continue
            break
        req.encoding = "gb2312"
        text = req.text
        html = etree.HTML(text)
        name = html.xpath('//div[@class="el"]/span/a/text()')
        for i in name:
            if i not in self.namelist and i[-4:] == '有限公司':
                self.namelist.append(i)
        
    def ynames(self,namelist):
        names = []
        for i in namelist:
            names.append(i)
            if len(names) == 2000:
                yield names
                names = []
        else:
            if names != []:
                yield names

    def write(self,namelist,addres,hangye):
        for names in self.ynames(namelist):
            time.sleep(1)
            now = time.strftime('%m%d%H%M%S', time.localtime(time.time()))
            book = xlwt.Workbook()
            sheet = book.add_sheet('my_book')
            sheet.write(0, 0, '公司名称')
            for index in range(len(names)):
                sheet.write(index+1,0,names[index])
            if os.path.exists('./爬取好的'):
                pass    
            else:
                os.mkdir('./爬取好的')
            book.save('./爬取好的/' + '51job' + addres + '_' + hangye + now + '.xls')
            
    def main(self):
        addres,hangye,cookie,page,url = self.search()
        self.driver.quit()
        print('请等待...')
        self.put_queue(page,url)
        t1 = time.time()
        for i in range(page):
            url = self.queue.get()
            t = threading.Thread(target=self.request,args=(cookie,Job.headers,url))
            t.start()
            self.th_size.append(t)
        for i in self.th_size:
            i.join()
        t2 = time.time() 
        print('总共抓取' + str(len(self.namelist)) +'家企业')
        print('耗时:' + str(t2-t1) + '秒')
        self.write(self.namelist,addres,hangye)
        
if __name__ == "__main__":
    job = Job()
    job.main()


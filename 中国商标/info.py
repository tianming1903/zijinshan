from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
import threading 
import queue
import time,xlrd,xlwt
import sys ,os
import requests
from requests.exceptions import Timeout
import json
from json.decoder import JSONDecodeError

class SB_info(object):
    def __init__(self,url):
        self.status = 0
        self.link = url
        self.queue = queue.Queue()
        opt = webdriver.ChromeOptions()
        opt.set_headless()
        self.driver = webdriver.Chrome(options=opt)
        self.driver.get(self.link)
        self.url = 'http://sbgg.saic.gov.cn:9080/tmann/annInfoView/annSearchDG.html'
        self.cookie = {}
        # 这是可能要更换的
        self.header = {
            "Accept":"application/json, text/javascript, */*; q=0.01",
            "Accept-Encoding": "gzip, deflate",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
            "Connection": "keep-alive",
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "Host": "sbgg.saic.gov.cn:9080",
            "Origin": "http://sbgg.saic.gov.cn:9080",
            "User-Agent": "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.131 Safari/537.36",
            "X-Requested-With": "XMLHttpRequest",
        }
        self.data = {
            "page": '',
            "rows": 20,
            "annNum": '',
            "annType":"",
            "tmType": "",
            "coowner": "",
            "recUserName": "",
            "allowUserName": "",
            "byAllowUserName": "",
            "appId": "",
            "appIdZhiquan": "",
            "bfchangedAgengedName": "", 
            "changeLastName": "",
            "transferUserName": "",
            "acceptUserName": "",
            "regName": "",
            "tmName": "",
            "intCls": "",
            "fileType": "",
            "totalYOrN": "false",
            "appDateBegin": "",
            "appDateEnd": "",
            "agentName": ""
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

    def clicks(self):
        # 点击查看到图片页面
        try:
             WebDriverWait(self.driver,20).until(EC.visibility_of_all_elements_located((By.XPATH,'//tbody/tr/td/a')))
        except TimeoutException:
            sys.exit('连接超时，程序被迫退出')

        # 获取总页数
        pages = self.driver.find_element_by_xpath('//div[@id="pages"]//td[6]/span').text
        page = ''
        for i in pages:
            if ord(i) in [x for x in range(48,58)]:
                page += i
        print('总页数是:' + page)

        # 获取cookie
        self.driver.find_element_by_xpath('//div[@id="pages"]//td[8]').click()
        try:
             WebDriverWait(self.driver,20).until(EC.visibility_of_all_elements_located((By.XPATH,'//tbody/tr/td/a')))
        except TimeoutException:
            sys.exit('连接超时，程序被迫退出')
        for i in self.driver.get_cookies():
            self.cookie[i['name']] = i['value']
        return page

    def read_text(self,page):
        if os.path.exists('./info.txt'):
            with open('index.txt','r',encoding='utf-8') as f:
                l = list(int(i) for i in f)
                not_list = set(l) ^ set(list(range(1,int(page))))
                return not_list
        else:
            return list(range(1,int(page)))

    def create_thread(self,not_list,page):
        # 创建线程进行请求
        print('数据正在爬取....')
        for i in not_list:
            self.data['page'] = i
            self.data['annNum'] = self.link.split('=')[1]
            # 创建一个线程去执行数据的爬取操作
            t = threading.Thread(target=self.request,args=(i,))
            t.start()
            self.queue.put(t)
            time.sleep(1)
            if self.status == 1:
                print('等待处理完所有的线程')
                while not self.queue.empty():
                    self.queue.get().join()
                self.driver.refresh()
                l = self.clicks()
                self.status = 0
            elif i == int(page):
                while not self.queue.empty():
                    self.queue.get().join()
                print('数据爬取完毕')
        
    def request(self,num):
        try:
            req = requests.post(self.url,cookies=self.cookie,data=self.data,headers=self.header,timeout=30)
            info = json.loads(req.text)['rows']
        except Timeout:
            print('连接超时，此数据被丢弃')
            return
        except JSONDecodeError as e:
            self.status = 1
            print('错误是: ',e)
            return
        for i in info:
            g_qihao = i['ann_num']
            time = i['ann_date']
            g_type = i['ann_type']
            zhucehao = i['reg_num']
            shengqingren = i['regname']
            shangbiaoming = i['tm_name']
            try:
                with open('info.txt','a',encoding='utf-8') as f:
                    f.write(g_qihao + ',' + time + ',' + g_type + ',' + zhucehao + ',' + shengqingren + ',' + shangbiaoming)
                    f.write('\n')
            except TypeError:
                print('数据写入出现错误')
                continue
        print('爬取的页数是:' + str(num))
        with open('index.txt','a',encoding='utf-8') as f:
            f.write(str(num))
            f.write('\n')
                
    def write_excel(self):
        infos = []
        with open('info.txt','r',encoding='utf-8') as f:
            while True:
                try:
                    i = f.readline()
                except UnicodeDecodeError:
                    continue
                if len(infos) == 5000:
                    yield infos
                    infos = []
                if i == '':
                    yield infos
                    break
                infos.append(i)
                    
    def write(self):
        for infos in self.write_excel():
            qihao = infos[0].split(',')[0]
            time.sleep(1)
            now = time.strftime('%m%d%H%M%S', time.localtime(time.time()))
            book = xlwt.Workbook()
            sheet = book.add_sheet('my_book')
            sheet.write(0, 0, '公告期号')
            sheet.write(0, 1, '公告日期')
            sheet.write(0, 2, '公告类型')
            sheet.write(0, 3, '注册号')
            sheet.write(0, 4, '申请人')
            sheet.write(0, 5, '商标名称')
            x = 1
            for info in infos:
                info = info.split(',')
                for index in range(len(info)):
                    sheet.write(x,index,info[index])
                x += 1
            if os.path.exists('./' + qihao + '期商标信息'):
                pass
            else:
                os.mkdir('./' + qihao + '期商标信息')
            book.save('./' + qihao + '期商标信息/' + '商标注册' + now + '.xls')
        os.remove('info.txt')
        os.remove('index.txt')
        
    def main(self):
        page = self.clicks()
        not_list = self.read_text(page)
        self.create_thread(not_list,page)
        self.driver.quit()
        self.write()
        print('信息写入完毕')

if __name__ == "__main__":
    # 获取爬取的链接
    link = SB_info.read_excel()
    sb = SB_info(link)
    sb.main()

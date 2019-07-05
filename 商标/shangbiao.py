import requests
import xlwt
from lxml import etree
import time,sys
from selenium import webdriver
from requests.exceptions import Timeout
import os

class ShangBiao(object):
    def __init__(self):
        self.headers = {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3',
            'Accept-Encoding': 'gzip, deflate',
            'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
            'Cache-Control': 'max-age=0',
            'Connection': 'keep-alive',
            'Cookie': 'UM_distinctid=16b01b75be0dd-01ce4026c5adb6-3e385e0c-13c680-16b01b75be13b2; SANGFOR_AD_TMOAS=51947.7749.15402.0000; JSESSIONID=00000zw7GHpimS3zq-k0vOveb95:1bm10epmj; __jsluid=a42e52d1f771b529427878a887cda94c',
            'Host': 'wssq.saic.gov.cn:9080',
            'Upgrade-Insecure-Requests': '1',
            'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.131 Safari/537.36',
            }
        self.url = 'http://wssq.saic.gov.cn:9080/tmsve/agentInfo_getAgentDljg.xhtml'
        self.driver = webdriver.Chrome()
        self.driver.get(self.url)

    def request(self,number):
        # 请求数据页面链接获得post请求的一些数据
        if number == '1':
            sums = self.driver.find_elements_by_xpath('//center/font')[0].text
            filename = self.driver.find_element_by_xpath('//ul/a[1]').text
            tab = 0
            self.driver.quit()
        elif number == '0':
            self.driver.find_element_by_xpath('//ul/a[2]').click()
            filename = self.driver.find_element_by_xpath('//ul/a[2]').text
            time.sleep(1)
            sums = self.driver.find_elements_by_xpath('//center/font')[0].text
            tab = 1
            self.driver.quit()
        else:
            self.driver.quit()
            sys.exit('所选项不存在，请重新运行')

        namelist = []
        for i in range(1,int(int(sums)/30 + 1)):
            data = {
                "ifLogIn":'', 
                "agentBean.agentName":'',
                "tabValue": tab,
                "pagenum": i,
                "pagesize": 30,
                "sum": sums,
                "countpage": int(int(sums)/30 + 1),
                "gopage": i+1,
                "pagenum": i+1,
                "pagesize": 30,
                "sum": sums,
                "countpage": int(int(sums)/30 + 1),
                "gopage": i+1,
                }
            headres = {
                'Content-Length': '152',
                'Content-Type': 'application/x-www-form-urlencoded',
                'Origin': 'http://wssq.saic.gov.cn:9080',
                'Referer': 'http://wssq.saic.gov.cn:9080/tmsve/agentInfo_getAgentDljg.xhtml'
            }
            headres = dict(self.headers,**headres)
            try:
                req = requests.post(self.url,data=data,headers=headres,timeout=10)
            except Timeout:
                continue  
            req.encoding = 'utf-8'
            html = etree.HTML(req.text)
            llist = html.xpath('//table[@class="import_tab"]/tr[@height="15"]')
            print(i,len(llist))
            for x in llist:
                name = x.xpath('./td/a/text()')[0]
                namelist.append(name.strip())
        return (namelist,filename)

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

    def write_excel(self,namelist,filename):
        for names in self.ynames(namelist):
            time.sleep(1)
            now = time.strftime('%m%d%H%M%S', time.localtime(time.time()))
            book = xlwt.Workbook()
            sheet = book.add_sheet('my_book')
            sheet.write(0, 0, '公司名称')
            for index in range(len(names)):
                sheet.write(index+1,0,names[index])
            if os.path.exists('爬取好的'):
                pass
            else:
               os.mkdir('./爬取好的')
            book.save('./爬取好的/' + filename + now + '.xls')
                    
if __name__ == "__main__":
    print('爬取的网站地址是:')
    print('http://wssq.saic.gov.cn:9080/tmsve/agentInfo_getAgentDljg.xhtml')
    number = input('爬取代理机构名单(不包含律师按1,包含律师按0):')
    sb = ShangBiao()
    names,filename = sb.request(number)
    sb.write_excel(names,filename)
    print('爬取完毕')

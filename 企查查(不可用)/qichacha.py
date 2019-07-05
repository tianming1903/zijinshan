# -*- coding: utf-8 -*-
import time
from Login import Login
from selenium.common.exceptions import NoSuchElementException,TimeoutException
import xlwt,xlrd
import os
import random

class TianyanSpider():
    def exists(self):
        #判断存储信息文件是否存在
        index = 0
        if os.path.exists('info.txt'):
            with open('info.txt','r',encoding='utf-8') as f:
                for x in f:
                    index = x.split(',')[0]
                return int(index)+1
        else:
            return 1

    def Identification(self):
        # 获取要读取文件的名称
        file = next(os.walk('./'))
        for i in file[2]:
            if ('.xlsx' in i) or ('.xls' in i):
                return i
        else:
            exit('添加要爬取的excel名单')

    def read(self,number,filename):
        # 读取excel的公司,
        '''
        excel文件的格式，只有名单这一列，而且必须从第二行开始
        '''
        data = xlrd.open_workbook(filename)
        sheet = data.sheets()[0]
        names = sheet.col_values(0)[number:]
        return names

    def login(self):
        object = Login().start()
        return object

    def search(self,object,names,number):
        # 按公司名进行搜索
        time.sleep(1)
        object.get('https://www.qichacha.com/firm_1d8c67930434ab1b1f87a2b7a203ad10.shtml')
        time.sleep(2)
        for name in names:
            object.find_element_by_xpath('//div[@class="input-group"]/input[1]').clear()
            name = name.strip()
            time.sleep(1)
            object.find_element_by_xpath('//div[@class="input-group"]/input[1]').send_keys(name)
            time.sleep(random.randint(1,2))
            try:
                object.find_element_by_xpath('//section[@id="header-search-list"]//a[1]').click()
            except (NoSuchElementException,IndexError,TimeoutException):
                number += 1
                continue
            self.extract(object,number)
            number += 1

    def extract(self,object,num):
        # 进行数据的提取
        time.sleep(2)
        textlist = [str(num)]
        # 公司名称
        try:
            name = object.find_element_by_xpath('//div[@class="content"]//h1').text
        except NoSuchElementException:
            return

        textlist.append(name)
        # 运营状态
        try:
            zhuangtai = object.find_element_by_xpath('//div[@class="row tags"]//span[1]').text
        except NoSuchElementException:
            return
        if zhuangtai != "存续":
            return
            
        try:
            if  '基本信息' in object.find_element_by_xpath('//div[contains(@class,"company-nav")]/div[2]/a').text:
                object.find_element_by_xpath('//div[contains(@class,"company-nav")]/div[2]/a').click()
            # 注册时间
            Established = object.find_element_by_xpath('//section[contains(@class,"panel b-a")]/table[@class="ntable"][2]//tr[2]/td[4]').text
            # 注册资本
            capital = object.find_element_by_xpath('//section[contains(@class,"panel b-a")]/table[@class="ntable"][2]//tr[1]/td[2]').text
            # 统一信用代码
            code = object.find_element_by_xpath('//section[contains(@class,"panel b-a")]/table[@class="ntable"][2]//tr[3]/td[2]').text
            textlist.extend([Established,capital,code])
        except (NoSuchElementException,IndexError):
            return
        # 法人
        try:
            Representative = object.find_element_by_xpath('//div[@class="pull-left"]//h2').text
        except NoSuchElementException:
            return
        textlist.append(Representative)
        # 电话
        phone = ''
        try:
            phone = object.find_element_by_xpath('//div[@class="dcontent"]/div[1]//span[@class="cvlu"]/span').text
        except NoSuchElementException:
            textlist.append('暂无')
        else:
            textlist.append(phone)
        # 点击知识产权加载细分
        # 判断是否为上市公司
        try:
            if  '基本信息' in object.find_element_by_xpath('//div[contains(@class,"company-nav")]/div[2]/a').text:
                object.find_element_by_xpath('//div[contains(@class,"company-nav")]/div[7]/a').click()
                time.sleep(1)
            else:
                object.find_element_by_xpath('//div[contains(@class,"company-nav")]/div[6]/a').click()
                time.sleep(1)
            l = object.find_element_by_xpath('//div[@class="data_div"][4]//div[@class="panel-body"]/a[1]').text.split(' ')[1]
            textlist.append(l)
            l = object.find_element_by_xpath('//div[@class="data_div"][4]//div[@class="panel-body"]/a[2]').text.split(' ')[1]
            textlist.append(l)
            l = object.find_element_by_xpath('//div[@class="data_div"][4]//div[@class="panel-body"]/a[5]').text.split(' ')[1]
            textlist.append(l)
            l = object.find_element_by_xpath('//div[@class="data_div"][4]//div[@class="panel-body"]/a[4]').text.split(' ')[1]
            textlist.append(l)
            l = object.find_element_by_xpath('//div[@class="data_div"][4]//div[@class="panel-body"]/a[6]').text.split(' ')[1]
            textlist.append(l)
            l = object.find_element_by_xpath('//div[@class="data_div"][4]//div[@class="panel-body"]/a[3]').text.split(' ')[1]
            textlist.append(l)
        except NoSuchElementException:
            textlist.extend(['0','0','0','0','0','0'])
        # 邮箱
        email = object.find_elements_by_xpath('//div[@class="dcontent"]/div[2]//span[@class="cvlu"]')[0].text
        textlist.append(email)
        # 官网
        try:
            url = object.find_element_by_xpath('//div[@class="dcontent"]/div[1]/span/a').get_attribute('href')
        except NoSuchElementException:
            textlist.append('暂无')
        else:
            textlist.append(url)
        # 公司地址
        try:
            address = object.find_elements_by_xpath('//div[@class="dcontent"]/div[2]//span[@class="cvlu"]/a[1]')
        except NoSuchElementException:
            return
        address = address[-1].get_attribute('onclick').split("'")[1]
        textlist.append(address)
        
        if (textlist[6] == '暂无') and (textlist[-2] == '暂无') and (textlist[-3] == '暂无'):
            return
        
        with open('info.txt', 'a', encoding='utf-8') as f:
            f.write((',').join(textlist))
            f.write('\n')

    def readinfo(self,filename = 'info.txt'):
        text = []
        with open (filename,'r',encoding='utf-8') as f:
            for i in f:
                info = i.split(',')[1:]
                if info in text:
                    continue
                text.append(info)
                if len(text) ==2000:
                    yield text
                    text = []
            else:
                if text == []:
                    return
                yield text

    def write(self,name):
        #把信息写入到excel
        for text in self.readinfo():
            time.sleep(1)
            now = time.strftime('%m%d%H%M%S', time.localtime(time.time()))
            book = xlwt.Workbook()
            sheet = book.add_sheet('my_book')
            sheet.write(0, 0, '公司名称')
            sheet.write(0, 1, '注册时间')
            sheet.write(0, 2, '注册资本')
            sheet.write(0, 3, '统一信用代码')
            sheet.write(0, 4, '法人')
            sheet.write(0, 5, '电话')
            sheet.write(0, 6, '商标信息')
            sheet.write(0, 7, '专利信息')
            sheet.write(0, 8, '软件著作权')
            sheet.write(0, 9, '作品著作权')
            sheet.write(0, 10, '网站备案')
            sheet.write(0, 11,  '资质')
            sheet.write(0, 12,  '邮箱')
            sheet.write(0, 13, '网址')
            sheet.write(0, 14, '地址')
            x = 1
            for infos in text:
                for num in range(len(infos)):
                    sheet.write(x,num,infos[num])
                x += 1
            book.save('info/' + name.split('.')[0] + '_' + now + '.xls')
        os.remove('info.txt')

    def main(self):
        index = self.exists()
        name = self.Identification()
        names = self.read(index,name)
        object = self.login()
        self.search(object,names,index)
        self.write(name)

if __name__ == "__main__":
    spider = TianyanSpider()
    spider.main()

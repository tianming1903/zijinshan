# -*- coding: utf-8 -*-
import time,random
from Login import Login
from selenium.common.exceptions import NoSuchElementException,StaleElementReferenceException
import xlwt,xlrd
import os,sys

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
            sys.exit('添加要爬取的excel名单')
    
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
    
    def search(self,object,names,number,status):
        # 按公司名进行搜索
        for name in names:
            name = name.strip()
            object.find_elements_by_xpath('//input[@type="search"]')[0].send_keys(name)
            time.sleep(random.randint(1,2))
            try:
                object.find_elements_by_xpath('//div[@class="name js-text"]')[0].click()
            except (NoSuchElementException,IndexError):
                object.find_elements_by_xpath('//input[@type="search"]')[0].clear()
                number += 1
                continue
            except StaleElementReferenceException:
                object.quit()
                sys.exit('程序被迫中止，请重新启动')
            self.fanhui(object,number,status)
            number += 1
    
    def fanhui(self,object,num,status):
        # 切换到新的页面进行数据的提取
        window = object.window_handles
        object.switch_to.window(window[1])
        self.extract(object,num,status)
        # 关闭新页面，切换到原页面输入下一个公司的名称进行查询
        object.close()
        object.switch_to.window(window[0])
        object.find_elements_by_xpath('//input[@type="search"]')[0].clear()
    
    def extract(self,object,num,status):
        # 进行数据的提取
        time.sleep(1)
        textlist = [str(num)]
        while True:
            try:
                name = object.find_element_by_xpath('//div[@class="header"]/h1').text
            except NoSuchElementException:
                try:
                    object.find_element_by_xpath('//div[@class="container -body"]')
                except NoSuchElementException:
                    time.sleep(5)
                    continue
                return
            textlist.append(name)
            break
    
        # 是否对公司名称进行过滤
        if (status.strip()).lower() == 'y':
            if name[-4:] != '有限公司':
                return
    
        # 获取运营状态
        try:
            zhuangtai = object.find_element_by_xpath('//div[@class="tag-list"]/div[1]').text
        except NoSuchElementException:
            return
        if zhuangtai != '存续' and zhuangtai != '在业':
            return
            
        # 注册时间，注册资本，实缴资本，信用代码,行业
        try:
            info = object.find_element_by_xpath('//div[@class="data-content"]/table[2]')
            Established = info.find_element_by_xpath('.//tr[2]//div').text
            textlist.append(Established)
            capital = info.find_element_by_xpath('.//tr[1]//div').text
            textlist.append(capital)
            pin_capital = info.find_element_by_xpath('.//tr[1]/td[4]').text
            textlist.append(pin_capital)
            code = info.find_element_by_xpath('.//tr[3]/td[2]').text
            textlist.append(code)
            industry = info.find_element_by_xpath('.//tr[5]/td[4]').text
            textlist.append(industry)
        except (IndexError,NoSuchElementException):
            return
    
        # 法人
        try:
            Representative = object.find_element_by_xpath('//div[@class="name"]/a').text
            textlist.append(Representative)
        except NoSuchElementException:
            return
    
        # 电话
        photo = object.find_element_by_xpath('//div[contains(@class,"sup-ie-company-header-child-1")]/span[2]').text
        if len(photo) == 11:
            textlist.append('/')
            textlist.append(photo)
        else:
            textlist.append(photo)
            textlist.append('/')
    
        # 知识产权，资质证书
        certificate = object.find_elements_by_xpath('//div[@id="nav-main-knowledgeProperty"]/div[@class="block-data"]/div/span')
        if certificate == []:
            textlist.extend(['0','0','0','0','0'])
        else:
            list1 = ['0','0','0','0','0']
            for number in range(len(certificate)):
                if certificate[number].text == "商标信息":
                    list1[0] = certificate[number + 1].text
                if certificate[number].text == "专利信息":
                    list1[1] = certificate[number + 1].text
                if certificate[number].text == "软件著作权":
                    list1[2] = certificate[number + 1].text
                if certificate[number].text == "作品著作权":
                    list1[3] = certificate[number + 1].text
                if certificate[number].text == "网站备案":
                    list1[4] = certificate[number + 1].text
            textlist.extend(list1)
        property = object.find_elements_by_xpath('//div[@id="nav-main-manageStatus"]/div[@class="block-data"]/div/span')
        if property == []:
            textlist.append(str(0))
        else:
            for number in range(len(property)):
                if property[number].text == "资质证书":
                    textlist.append(property[number + 1].text)
                    break
            else:
                textlist.append(str(0))
        
        # 邮箱
        email = object.find_element_by_xpath('//div[contains(@class,"sup-ie-company-header-child-2")]/span[2]').text
        textlist.append(email)
    
        # 网址
        try:
            url = object.find_element_by_xpath('//div[contains(@class,"sup-ie-company-header-child-1")]//a').get_attribute('href')
        except NoSuchElementException:
            textlist.append('暂无信息')
        else:
            textlist.append(url)
            
        # 地址
        try:
            address = object.find_element_by_xpath('//div[contains(@class,"sup-ie-company-header-child-2")]/div/div').text
        except NoSuchElementException:
            textlist.append('暂无信息')
        else:
            textlist.append(address)
    
        # 电话 邮箱 网站必须要有一个，否则不要
        if (textlist[7] == '暂无信息') and (textlist[-2] == '暂无信息') and (textlist[-3] == '暂无信息'):
            return
    
        with open('info.txt', 'a', encoding='utf-8') as f:
            f.write((',').join(textlist))
            f.write('\n')
    
    def write(self,name):
        #把信息写入到excel
        text = []
        with open ('info.txt','r',encoding='utf-8') as f:
            for i in f:
                info = i.split(',')[1:]
                if info in text:
                    continue
                text.append(info)
        now = time.strftime('%m%d%H%M%S', time.localtime(time.time()))
        book = xlwt.Workbook()
        sheet = book.add_sheet('my_book')
        sheet.write(0, 0, '公司名称')
        sheet.write(0, 1, '注册时间')
        sheet.write(0, 2, '注册资本')
        sheet.write(0, 3, '实缴资本')
        sheet.write(0, 4, '统一信用代码')
        sheet.write(0, 5, '行业')
        sheet.write(0, 6, '法人')
        sheet.write(0, 7, '电话')
        sheet.write(0, 8, '手机号码')
        sheet.write(0, 9, '商标信息')
        sheet.write(0, 10, '专利信息')
        sheet.write(0, 11, '软件著作权')
        sheet.write(0, 12, '作品著作权')
        sheet.write(0, 13, '网站备案')
        sheet.write(0, 14,  '资质')
        sheet.write(0, 15,  '邮箱')
        sheet.write(0, 16, '网址')
        sheet.write(0, 17, '地址')
        x = 1
        for infos in text:
            for num in range(len(infos)):
                sheet.write(x,num,infos[num])
            x += 1
        if os.path.exists('./爬取好的'):
           pass
        else:
            os.mkdir('爬取好的')
        book.save('./爬取好的/' + name.split('.')[0] + '_' + now + '.xls')
        os.remove('info.txt')
    
    def main(self):
        status = input('是否选择只保留有限公司(y/n):')
        index = self.exists()
        name = self.Identification()
        names = self.read(index,name)
        object = self.login()
        self.search(object,names,index,status)
        self.write(name)

if __name__ == "__main__":
    spider = TianyanSpider()
    spider.main()
    print('名单爬取完毕')


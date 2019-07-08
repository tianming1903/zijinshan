# -*- coding: utf-8 -*-
import time, random
from login import Login
import xlrd,os,xlwt
from selenium.common.exceptions import NoSuchElementException

class TianyanSpider():
    def exists(self):
        if os.path.exists('links.txt'):
            with open('links.txt','r',encoding='utf-8') as f:
                l = len(f.readlines())
                if '\n' in f.readlines():
                    return l
                return l + 1 
        else:
            return 1

    def mingdan(self):
        files = next(os.walk('./'))
        for i in files[2]:
            if ('.xls' in i) or ('xlsx' in i):
                return i

    def read(self,filename,number):
        # 读取links表获取要爬取的链接
        book = xlrd.open_workbook(filename)
        sheet = book.sheets()[0]
        links = sheet.col_values(0)[number:]
        driver = Login().main()
        return links,driver

    def parse(self,links,driver):
        for link in links:
            nameslist = []
            driver.get(link.strip())
            time.sleep(3)
            # 获取地址名称作为文件名
            name1 = driver.find_element_by_xpath('//div[@class="filter-scope"]/a[1]').text.split('：')[1]
            name2 = driver.find_element_by_xpath('//div[@class="filter-scope"]/a[2]').text.split('：')[1]
            filename = name1 + '_' + name2
            while True:
                # 匹配信息
                items = driver.find_elements_by_xpath('//div[contains(@class,"search-item")]')
                for info in items:
                    try:
                        name = info.find_element_by_xpath('.//div[@class= "header"]/a').text
                        person = info.find_element_by_xpath('.//div[contains(@class,"row ")]/div[1]/a').text
                        capital = info.find_element_by_xpath('.//div[contains(@class,"row ")]/div[2]/span').text
                        zctime = info.find_element_by_xpath('.//div[contains(@class,"row ")]/div[3]/span').text
                    except NoSuchElementException:
                        continue
                    try:
                        tel = info.find_element_by_xpath('.//div[contains(@class,"col")][1]/span/span[1]').text
                        email = info.find_element_by_xpath('.//div[contains(@class,"col")][2]/span[2]').text
                    except NoSuchElementException:
                        continue
                    info = name + ',' + person + ',' + capital + ',' + zctime + ',' + tel + ',' + email
                    nameslist.append(info)
                try:
                    next_page = driver.find_element_by_xpath('//a[@class="num -next"]')
                except NoSuchElementException:
                    self.write(nameslist,filename)
                    with open('links.txt','a',encoding='utf-8') as f:
                        f.write(link.strip())
                        f.write('\n')
                    break
                driver.execute_script("arguments[0].click();",next_page)
                # 进入验证
                time.sleep(random.randint(2,3))
                while True:
                    try:
                        driver.find_element_by_xpath('//div[@class="my_btn_web"and @id="submitie"]')
                    except NoSuchElementException:
                        break
                    time.sleep(3)
                    continue
        os.remove('./links.txt')
        driver.quit()
    
    def readinfo(self,namelist):
        text = []
        for i in namelist:
            text.append(i)
            if len(text) == 2000:
                yield text
                text = []
        else:
            if text != []:
                yield text

    def write(self,namelist,filename):
        # 把信息写入到excel
        for text in self.readinfo(namelist):
            time.sleep(1)
            now = time.strftime('%m%d%H%M%S', time.localtime(time.time()))
            book = xlwt.Workbook()
            sheet = book.add_sheet('my_book')
            sheet.write(0, 0, '公司名称')
            sheet.write(0, 1, '法人')
            sheet.write(0, 2, '注册资本')
            sheet.write(0, 3, '注册时间')
            sheet.write(0, 4, '电话')
            sheet.write(0, 5, '邮箱')
            x = 1
            for infos in text:
                ilist = infos.strip().split(',')
                for i in range(len(ilist)):
                    sheet.write(x, i, ilist[i])
                x += 1
            if os.path.exists('./爬取好的'):
                pass
            else:
                os.mkdir('./爬取好的')
            book.save('./爬取好的/' + filename + now + '.xls')
        print('爬取和保存完一条链接,开启下一条的抓取')
        
if __name__ =="__main__":
    ty = TianyanSpider()
    number = ty.exists()
    filename = ty.mingdan()
    links,driver = ty.read(filename,number)
    ty.parse(links,driver)
    print('所有链接爬取完毕')

import time,xlwt
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException,ElementClickInterceptedException,TimeoutException
import os

class Remingwang(object):
    def __init__(self):
        self.driver = webdriver.Chrome()
        # 次数可更改爬取的链接
        self.driver.get('http://www.chinahbjob.com/Job/Search.aspx')
    
    def request(self):
        nameslist = []
        x = 0
        while True:
            names = self.driver.find_elements_by_xpath('//li/div/a/span')
            if names == []:
                break
            for i in names:
                if i.text[-4:] == '有限公司' and i.text not in nameslist:
                    nameslist.append(i.text)
            # 获取下一页按钮
            button = self.driver.find_elements_by_xpath('//div[@class="pageNav"]/a')
            nexts = self.driver.find_element_by_xpath('//div[@class="pageNav"]/a['+ str(len(button)) + ']')
            self.driver.execute_script("arguments[0].click();",nexts)
            x += 1
            print('当前爬取页数: ' + str(x))
        self.driver.quit()
        return nameslist
    
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

    def write(self,namelist):
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
            book.save('./爬取好的/' + './湖北人才网' + now + '.xls')
                    
if __name__ == "__main__":
    rmw = Remingwang()
    namelist = rmw.request()
    rmw.write(namelist)
    print('数据爬取完毕')
    print('总共抓取' + str(len(namelist)) + '条数据')

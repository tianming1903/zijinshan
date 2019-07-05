from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
import time
import xlwt
import os

class ZhiLian(object):
    def __init__(self):
        self.driver = webdriver.Chrome()
        self.driver.maximize_window()
        self.driver.get('https://sou.zhaopin.com/?jl=489&sf=0&st=0')
    
    def search(self):
        input('请按任意键确定选择完成:')
        diqu = self.driver.find_element_by_xpath('//span[contains(@class,"current-city__name")]').text
        try:
            hangye = self.driver.find_element_by_xpath('//a[contains(@class,"query-industry__subIndustry--active")]').text
            hangye = ('').join(hangye.split('/'))
        except NoSuchElementException:
            hangye = ''
        filename = diqu + '_' + hangye
        #获取公司名单
        namelist = []
        sign = 1
        while True:
            names = self.driver.find_elements_by_xpath('//div[contains(@class,"commpanyName")]/a')
            for name in names:
                name = name.text.strip()
                if (name not in namelist) and (name[-4:] == '有限公司'):
                    namelist.append(name)
            # 下拉滚动条到底部
            js = "var q=document.documentElement.scrollTop=100000"
            self.driver.execute_script(js)
            # 点击翻页加载新的页面
            try:
                button = self.driver.find_elements_by_xpath('//button[@class="btn soupager__btn"]')
            except NoSuchElementException:
                break
            if len(button) == 2:
                button[1].click()
            elif len(button) == 1 and sign == 1:
                button[0].click()
                sign = 0
            else:
                break
            time.sleep(5)
        self.driver.quit()
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

    def write(self,namelist,filename):
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
            book.save('./爬取好的/' + './智联' + filename + now + '.xls')

if __name__ == "__main__":
    t1 = time.time()
    zl = ZhiLian()
    namelist,filename  = zl.search()
    zl.write(namelist,filename)
    t2 = time.time()
    print('名单抓取完毕')
    print('总共抓取信息'+ str(len(namelist)) + '条')
    print('耗时' + str(t2-t1) + '秒')
    

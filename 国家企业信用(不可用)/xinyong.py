from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException,NoSuchElementException,ElementClickInterceptedException
import sys,os,random,re,time
import xlwt

class xinyong(object):
    def __init__(self):
        self.driver = webdriver.Chrome()
        self.start_url = 'http://www.gsxt.gov.cn/corp-query-entprise-info-xxgg-100000.html'
        self.serial = [2,3,4,6,7,17,18,20,21,22,23]

    def request(self,i):
        self.driver.get(self.start_url)
        try:
            WebDriverWait(self.driver,10).until(EC.element_to_be_clickable((By.CLASS_NAME,'page_left')))
        except TimeoutException:
            self.driver.quit()
            sys.exit('页面没有加载完成而失败，请重新运行')

        # 点击不同的栏目
        classification = self.driver.find_element_by_xpath('//div[@class="page_left"]/a[' + str(i) + ']')
        classification.click()
        # 等待加载
        while True:
            try:
                WebDriverWait(self.driver,30).until(EC.presence_of_element_located((By.XPATH,'//div/div[@style="display: none;"]')))
            except TimeoutException:
                classification.click()
                continue
            break
        # 产生一个文件夹
        os.mkdir(classification.text)
        # 点击不同的城市和省份
        province = self.driver.find_elements_by_xpath('//div[@class="label-list"]/div/label')
        for x in province:
            self.driver.execute_script('arguments[0].click();',x)
            # 等待翻页加载出来
            while True:
                try:
                    WebDriverWait(self.driver,30).until(EC.element_to_be_clickable((By.XPATH,'//li[contains(@class,"paginate_button next")]/a')))   
                except TimeoutException:
                    x.click()
                    continue
                break
            # 生成文件的名称
            filename = classification.text + '_' + x.text
            # 获取页码数量并点击
            pages = self.driver.find_elements_by_xpath('//ul[@class="pagination"]/li')
            namelist = []
            if len(pages) > 4:
                for y in range(3,len(pages)-1):
                    button = self.driver.find_element_by_xpath('//ul[@class="pagination"]/li[' + str(y) + ']/a')
                    self.driver.execute_script('arguments[0].click();',button)
                    # 等待加载
                    while True:
                        try:
                            WebDriverWait(self.driver,30).until(EC.presence_of_element_located((By.XPATH,'//div/div[@style="display: none;"]')))
                        except TimeoutException:
                            button.click()
                            continue
                        break
                    # 直接提取公司的名称就可以了
                    names = self.driver.find_elements_by_xpath('//tbody/tr[@role="row"]//a')
                    for name in names:
                        name = name.text
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
                        elif '营业执照' in name:
                            index = name.index('移')
                            name = name[:index]
                        if name not in namelist and len(name) >= 5:
                            namelist.append(name)

                book = xlwt.Workbook()
                table = book.add_sheet('my_book')
                table.write(0,0,'公司名称')
                for i in range(1,len(namelist)+1):
                    table.write(i,0,namelist[i-1])
                book.save(classification.text + '/' + filename + '_' + '.xls')

    def main(self):
        for i in self.serial:
            self.request(i)
        self.driver.quit()

if __name__ == "__main__":
    xy = xinyong()
    xy.main()
    print('爬取完成')


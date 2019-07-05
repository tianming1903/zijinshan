from selenium import webdriver
import time
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException,TimeoutException,ElementClickInterceptedException
import xlrd,sys,os
import time
from selenium.webdriver.chrome.options import Options

class Boss():
    def __init__(self):
        self.options = Options()
        self.options.add_experimental_option('excludeSwitches', ['enable-automation'])
        self.time = time.time()
        self.demand = {}
        self.recruiter = {}

    # 获取所有招聘者名单
    def obtain_file (self):
        filenames = next(os.walk('./'))[2]
        x = 0
        for i in filenames:
            if ('.xls' in i) or ('xlsx' in i):
                self.recruiter[(i.split('.')[0])] = x
                x += 1
        print('可供选择的招聘者:{}'.format(self.recruiter))
        while True:
            try:
                num = int(input('请选择你想招聘的人(选择对应的数字即可):'))
            except ValueError:
                print('输入的值不是数字，请重新输入')
                continue
            if num not in self.recruiter.values():
                print('你输入的值不是我想要的，请重新输入')
                continue
            break
        # 返回招聘者名字
        name = list(self.recruiter.keys())[list(self.recruiter.values()).index(num)]
        return name
        
    # 根据所选择的招聘者获取招聘详细信息
    def huoqu(self,name):
        # 根据名字找到文件
        filenames = next(os.walk('./'))[2]
        for i in filenames:
            if name == i.split('.')[0]:
                break
        # 读取文件的信息信息
        book = xlrd.open_workbook(i)
        sheet = book.sheets()[0]
        name_password = sheet.row_values(0)
        post = sheet.col_values(1)[1:]
        number = sheet.col_values(2)[1:]
        for demand,num in zip(post,number):
            if demand == '' or num == '':
                continue
            self.demand[int(demand)] = int(num)
        return name_password

    def login(self,name_password):
        while True:
            num = input('输入登录方式(密码登录按1，扫码登录按0):')
            if num != '0' and num != '1':
                print('选择的登录方式不存在，请重新选择')
                continue
            break

        # 开始进行登录
        driver = webdriver.Chrome(options=self.options)
        driver.get('https://www.zhipin.com/chat/im?mu=recommend')
        driver.maximize_window()

        if num == "0":
            # --------采用扫码登录---------
            driver.find_elements_by_xpath('//div/span[@class="link-scan"]')[0].click()
            while True:
                try:
                    driver.find_element_by_xpath('//a[@ka="menu-geek-recommend"]')
                except NoSuchElementException:
                    time.sleep(3)
                    continue
                break
        elif num == "1":
        # ------采用账号密码---------
            while True:
                time.sleep(10)
                if driver.find_elements_by_class_name("nc-lang-cnt")[0].text == "验证通过":
                    break
                else:
                    driver.refresh()
            driver.find_elements_by_xpath('//input[@type="tel"]')[0].send_keys(str(name_password[0]).split('.')[0])
            time.sleep(1)
            driver.find_element_by_xpath('//input[@type="password"]').send_keys(name_password[1])
            time.sleep(2)
            driver.find_element_by_xpath('//div[@class="row-code nc-container"]/../following-sibling::div[1]/button').click() 
        return driver

    def dazhaohu(self,driver):
        namelist = []
        try:
            WebDriverWait(driver,5).until(EC.element_to_be_clickable((By.CLASS_NAME,"menu-recommend ")))
        except TimeoutException:
            return
        driver.find_element_by_xpath('//a[@ka="menu-geek-recommend"]').click()
        for num in self.demand:
            # 选择岗位
            time.sleep(1)
            i = 1
            while True:
                try:
                    button = driver.find_elements_by_xpath('//div[@class="dropdown-wrap"]/span')[1]
                except NoSuchElementException:
                    time.sleep(1)
                    continue
                driver.execute_script("arguments[0].click();",button)
                break
            time.sleep(1)
            driver.find_elements_by_xpath('//div[@class="dropdown-wrap dropdown-menu-open"]//li')[num].click()
            
            # 切换到iframe中去
            driver.switch_to_frame(driver.find_element_by_xpath('//iframe[@class="frame-container"]'))
            
            # 判断此岗位是否有人
            try:
                driver.find_element_by_xpath('//div[@class="data-tips"]')
            except NoSuchElementException:
                pass
            else:
                continue
            
            # 务必要加载出打招呼的人数
            while True:
                names = driver.find_elements_by_xpath('//div[@class="name"]')
                if 0 < len(names):
                    break
                time.sleep(3)
                continue
            time.sleep(2)
            
            # 开始打招呼
            while True:
                names = driver.find_elements_by_xpath('//div[@class="name"]')
                zhuangtai = driver.find_elements_by_xpath('//button')
                for x,y in zip(zhuangtai,names):
                    if x.text == "打招呼" and i <= self.demand[num]:
                        driver.execute_script("arguments[0].click();",x)
                        i += 1
                        namelist.append((y.text[:-4]).strip())
                        try:
                            driver.find_element_by_xpath('//div[@class="dialog-container"]')
                        except NoSuchElementException:
                            time.sleep(2)
                            continue
                        time.sleep(2)
                        driver.refresh()
                        print('打招呼人数:' + str(len(namelist)))
                        self.qiujianli(driver,namelist)
                        return
                while True:
                    try:
                        driver.find_element_by_xpath('//div[@class="loadmore"]/span') 
                    except NoSuchElementException:
                        time.sleep(2)
                        continue
                    break
                if '没有更多了' == driver.find_element_by_xpath('//div[@class="loadmore"]/span').text:
                    break
                if i < self.demand[num]:
                    js = "var q=document.documentElement.scrollTop=10000"
                    driver.execute_script(js)
                    while True:
                        if '...' in driver.find_element_by_xpath('//div[@class="loadmore"]/span').text:
                            continue
                        break
                    continue
                break

        # 退出iframe重选岗位
            driver.switch_to_default_content()
        print('打招呼人数:' + str(len(namelist)))
        self.qiujianli(driver,namelist)

    def qiujianli(self,driver,namelist):
        renshu = 0
        index = 0
        L = []
        # 点击沟通
        driver.find_element_by_xpath('//a[@ka="menu-im"]').click()
        while True:
            # 刷新所带来的问题解决
            while True:
                try:
                    WebDriverWait(driver,5).until(EC.element_to_be_clickable((By.XPATH,'//ul[@class="main-list"]/li')))
                except TimeoutException:
                    continue
                break
            time.sleep(60)
            m = driver.find_elements_by_xpath('//ul[@class="main-list"]/li')
            for i in range(1,len(m)+1):
                try:
                    driver.find_element_by_xpath('//ul[@class="main-list"]/li[' + str(i) + ']' + '/a/div[@class="figure"]/span')
                    name = driver.find_element_by_xpath('//ul[@class="main-list"]/li[' + str(i) + ']' + '/a//span[@class="name"]').text.strip()
                except NoSuchElementException:
                    continue
                if (name not in L) and (name in namelist):
                    driver.find_element_by_xpath('//ul[@class="main-list"]/li['+ str(i) + ']').click()
                    time.sleep(2)
                    driver.find_element_by_xpath('//div[@class="chat-editor"]//div/a[3]').click()
                    time.sleep(1)
                    driver.find_element_by_xpath('//div[@class="dialog-footer"]//span[2]').click()
                    L.append(name)
                    index += 1
            if renshu != index:
                renshu = index
                print('发出求简历人数' + str(index))
            driver.refresh()
            if int(self.time) + 3600 <= int(time.time()):
                break
        driver.quit()

    def main(self):
        name = self.obtain_file()
        name_password = self.huoqu(name)
        driver = boos.login(name_password)
        boos.dazhaohu(driver)
        print('招聘代码已经结束')

if __name__ == "__main__":  
    boos = Boss()
    boos.main()

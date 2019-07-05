import time,sys
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.common.exceptions import NoSuchElementException

class Login(object):
    def __init__(self):
        self.driver = webdriver.Chrome()
        self.name = '13797143525'
        self.password = 'li123456'

    def request(self):
        # 请求此页面进行登录
        self.driver.get('https://www.qichacha.com/')
        time.sleep(3)
        self.driver.find_element_by_xpath('//ul[@class="navi-nav pull-right"]/li[9]').click()
        time.sleep(1)
        self.driver.find_element_by_xpath('//div[@class="login-tab col3"][2]').click()

    def input(self):
        # 输入用户名和密码进行登录
        time.sleep(1)
        self.driver.find_element_by_xpath('//input[@id="nameNormal"]').send_keys(self.name)
        time.sleep(1)
        self.driver.find_element_by_xpath('//input[@id="pwdNormal"]').send_keys(self.password)
        time.sleep(1)
        # 拖动滑块
        yanzhengma = self.driver.find_element_by_xpath('//span[@class="nc_iconfont btn_slide"]')
        ActionChains(self.driver).click_and_hold(yanzhengma).perform()
        ActionChains(self.driver).move_by_offset(500,0).perform()
        ActionChains(self.driver).release().perform()
        time.sleep(1)
        try:
            self.driver.find_element_by_xpath('//div[@class="errloading"]')
        except NoSuchElementException:
            pass
        else:
            sys.exit('程序被迫退出')
        # 登录验证
        while True:
            try:
                self.driver.find_element_by_xpath('//div[@class="nc_scale_submit"]/span')
            except NoSuchElementException:
                time.sleep(2)
                self.driver.find_element_by_xpath('//form[@class="form-group login-form"]/div[3]/div/../following-sibling::button[1]').click()
                break
            while True:
                try:
                    self.driver.find_element_by_xpath('//div[@style="border-top-color: rgb(229, 229, 229);"]')
                except NoSuchElementException:
                    time.sleep(2)
                    continue
                time.sleep(2)
                self.driver.find_element_by_xpath('//form[@class="form-group login-form"]/div[3]/div/../following-sibling::button[1]').click()
                break
            break

    def start(self):
        self.request()
        self.input()
        return self.driver

if __name__ == "__main__":
        l = Login()
        l.start()

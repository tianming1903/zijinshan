import time
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException

class Login(object):
    def __init__(self):
        self.driver = webdriver.Chrome()
        self.name = '18571568742'
        self.password = 'cxr787868'

    def request(self):
        # 请求此页面进行登录
        self.driver.get('http://www.tianyancha.com/')
        time.sleep(3)
        self.driver.find_element_by_xpath('//div[@class="nav-item -home"]/a').click()
        time.sleep(1)
        self.driver.find_element_by_xpath('//div[@active-tab="1"]').click()

    def input(self):
        # 输入用户名和密码进行登录
        time.sleep(1)
        self.driver.find_elements_by_xpath('//input[@class="input contactphone"]')[2].send_keys(self.name)
        time.sleep(1)
        self.driver.find_element_by_xpath('//input[@class="input contactword input-pwd"]').send_keys(self.password)
        time.sleep(2)
        self.driver.find_element_by_xpath('//div[@tyc-event-ch="LoginPage.PasswordLogin.Login"]').click()

    def verification(self):
        while True:
            try:
                self.driver.find_element_by_xpath('//div[@class="gt_cut_fullbg gt_show"]')
            except NoSuchElementException:
                break
            else:
                time.sleep(3)
                continue

    def main(self):
        self.request()
        self.input()
        self.verification()
        return self.driver

if __name__ == "__main__":
        l = Login()
        l.main()

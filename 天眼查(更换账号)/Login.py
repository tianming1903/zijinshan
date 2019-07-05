import time,sys
from selenium import webdriver
from PIL import Image
from io import BytesIO
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from selenium.common.exceptions import NoSuchElementException,TimeoutException,StaleElementReferenceException
import os


class Login(object):
    def __init__(self):
        self.driver = webdriver.Chrome()
        self.driver.get('http://www.tianyancha.com/')
        self.name_pwd = self.read_txt()

    def read_txt(self):
        name_pwd = []
        # 读取账号名单
        filelist = next(os.walk('./'))[2]
        for i in filelist:
            if '.txt' in i and i != 'info.txt':
                with open(i,'r') as f:
                    for i in f:
                        dicts = {}
                        if i == '\n':
                            break
                        l = i.strip().split(',')
                        dicts[l[0]] = l[1]
                        name_pwd.append(dicts)
        return name_pwd

    def request(self):
        # 请求此页面进行登录
        while True:
            try:
                WebDriverWait(self.driver,10).until(EC.element_to_be_clickable((By.CLASS_NAME,'link-white')))
            except TimeoutException:
                sys.exit('网络较差,请稍后再试')
            except StaleElementReferenceException:
                self.driver.refresh()
                continue
            break
        self.driver.find_element_by_xpath('//div[@class="nav-item -home"]/a').click()
        time.sleep(1)
        try:
            self.driver.find_element_by_xpath('//div[@active-tab="1"]').click()
        except NoSuchElementException:
            sys.exit('网络较差,请稍后再试')

    def setphone(self,num):
        dic = self.name_pwd[num]
        for name,password in dic.items():
            name = name
            password = password
        return (name,password)
        
    def input(self,name,password):
        # 输入用户名和密码进行登录
        time.sleep(1)
        self.driver.find_elements_by_xpath('//input[@class="input contactphone"]')[2].send_keys(name)
        time.sleep(1)
        self.driver.find_element_by_xpath('//input[@class="input contactword input-pwd"]').send_keys(password)
        time.sleep(2)
        self.driver.find_element_by_xpath('//div[@tyc-event-ch="LoginPage.PasswordLogin.Login"]').click()
        
    def get_screenshot(self):
        # 获取整个屏幕的截图
        screenshot = self.driver.get_screenshot_as_png()
        srceenshot =Image.open(BytesIO(screenshot))
        return srceenshot
    
    # 获取验证码的位置
    def get_position(self):
        img = self.driver.find_element_by_xpath('//div[@class="gt_cut_bg gt_show"]')
        size = img.size
        location = img.location
        top,buttom,left,right = location['y'],location['y'] + size['height'],location['x'],location['x'] + size['width']
        return (top,buttom,left,right)

    # 截取验证码
    def get_geetest_image(self,filename = 'captcha1.png'):
        srceenshot = self.get_screenshot()
        top,buttom,left,right = self.get_position()
        captcha = srceenshot.crop((left,top,right,buttom))
        # captcha.save(filename)
        return captcha
    
    # 根据照片获取移动的距离
    def get_gap(self,image):
        left = 60
        # 遍历验证码的每个像素点,从滑块的右侧开始
        for x in range(left, image.size[0]):
            for y in range(image.size[1]):
                pixel1 = image.load()[x, y]
                if pixel1[0] <= 30 and pixel1[1] <= 100 and pixel1[2] <= 100:
                    return x
        else:
            return 50

    # 计算移动轨迹
    def get_track(self,Offset):
        track = []
        current = 0
        mid = Offset * 4 / 5
        t = 0.2
        v = 0
        while current < Offset:
            if current < mid:
                a = 20
            else:
                a = -40
            v0 = v
            v = v0 + a * t
            move = v0 * t + 1 / 2 * a * t * t
            current += move
            track.append(round(move,1))
        return track

    # 拖动滑块
    def move_to_gap(self,track):
        huakuai = self.driver.find_element_by_xpath('//div[@class="gt_slider"]/div[2]')
        ActionChains(self.driver).click_and_hold(huakuai).perform()
        for i in track:
            ActionChains(self.driver).move_by_offset(i,0).perform()
        time.sleep(0.5)
        ActionChains(self.driver).move_by_offset(-6,0).perform()
        time.sleep(0.5)
        ActionChains(self.driver).release().perform()

    def start(self,num):
        # 采用密码账号登录
        while True:
            self.request()
            name,pwd = self.setphone(num)
            # 输入密码和账号点击登录
            self.input(name,pwd)
            # 点击加载出第二张验证码
            try:
                WebDriverWait(self.driver,10).until(EC.presence_of_all_elements_located((By.CLASS_NAME,'gt_popup_box')))
            except TimeoutException:
                self.driver.refresh()
                continue
            while True:
                self.driver.maximize_window()
                self.driver.find_element_by_xpath('//div[@class="gt_slider"]/div[2]').click()
                # 等待验证码加载完成
                time.sleep(3)
                image = self.get_geetest_image('captcha2.png')
                # 获取偏移量
                offset = self.get_gap(image)
                # 计算出拖动的路径
                track = self.get_track(offset)
                # # 开始拖动
                self.move_to_gap(track)
                # 等待验证完成
                time.sleep(3)
                # 验证是否通过验证码
                try:
                    self.driver.find_element_by_xpath('//div[@class="gt_cut_fullbg gt_show"]')
                except NoSuchElementException:
                    self.driver.set_window_size(1000,800)
                    return self
                # 刷新页面
                if '出现错误' in self.driver.find_element_by_xpath('//div[@class="gt_info_text"]').text:
                    self.driver.refresh()
                    break
                self.driver.find_element_by_xpath('//a[@class="gt_refresh_button"]').click()
                # 等待验证码刷新完成
                time.sleep(3)

if __name__ == "__main__":
        l = Login()
        l.start(0)

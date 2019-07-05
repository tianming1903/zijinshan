from selenium import webdriver
import time
from PIL import Image
import pytesseract
from io import BytesIO

class Jiaofei(object):
    def __init__(self):
        self.driver = webdriver.Chrome()
        self.driver.maximize_window()
        self.driver.get('http://fee.cnipa.gov.cn/app/window/form_window_add.jsp?select-key:payment_type=3')
        self.js = 'var q=document.documentElement.scrollTop=10000'
        self.driver.execute_script(self.js)
        
    def read(self):
        pass

    def write(self):
        for i in range(1):
            infos = self.driver.find_elements_by_xpath('//ul[@id="ul_hkrxx"]/li//input')
            time.sleep(1)
            infos[2].send_keys('李天明')
            time.sleep(1)
            infos[3].send_keys('56')
            time.sleep(1)
            infos[4].send_keys('13797143525')
            time.sleep(1)
            self.driver.find_element_by_xpath('//input[@name="select-key:agency_id"]').click()
            time.sleep(1)
            browsebox = self.driver.find_elements_by_xpath('//div[@class="browsebox"]/div[@class="bs_l"]/p')
            for i in browsebox:
                if i.text == "国家知识产权局专利局":
                    i.click()
            time.sleep(1)
            self.driver.find_element_by_xpath('//input[@name="select-key:patent_type"]').click()
            time.sleep(1)
            browsebox = self.driver.find_elements_by_xpath('//div[@class="browsebox"]/div[@class="bs_l"]/p')
            for i in browsebox:
                if i.text == "国家申请/集成电路":
                    i.click()
            # 根据需求相应的增加
            for i in range(1):
                self.driver.find_element_by_xpath('//center/input[1]').click()
            infos = self.driver.find_elements_by_xpath('//tbody/tr')
            for i in infos:
                i.find_element_by_xpath('./td[3]//input').send_keys('452612892')
                time.sleep(1)
                i.find_element_by_xpath('./td[5]//div[2]/input').click()
                time.sleep(1)
                for x in i.find_elements_by_xpath('./td[5]//div[@class="bs_l"]/p'):
                    if x.text == "优先权要求费":
                        x.click()
                i.find_element_by_xpath('./td[6]//input[2]').send_keys('52')
        return self.driver

    def yanzheng(self,driver):
        time.sleep(2)
        i = 0
        while True:
            i += 1
            # 获取照片的链接
            img = driver.find_element_by_xpath('//div[@class="div_yzm"]/img[1]').get_attribute('src')
            # 打开一个新窗口并切换到新窗口
            # js = '"window.open(' + img + ');"'
            js = "window.open('" + img + "');"
            driver.execute_script(js)
            window = driver.window_handles
            driver.switch_to_window(window[1])
            # 获取整个屏幕的照片
            screenshot = driver.get_screenshot_as_png()
            screenshot = Image.open(BytesIO(screenshot))
            img1 = driver.find_element_by_xpath('//img')
            location = img1.location
            size = img1.size
            top,bottom,left,right = location['y'], location['y'] + size['height'], location['x'], location['x'] + size['width']
            # 截取验证码
            yanzhengma = screenshot.crop((left, top, right, bottom))
            yanzhengma.save('yanzhengma.png')
            # 关闭窗口切换到原来窗口
            driver.close()
            driver.switch_to_window(window[0])
            # 对照片进行底色处理
            image = Image.open('./yanzhengma.png')
            img2 = image.convert('RGBA')
            pixdata = img2.load()
            for y in range(img2.size[1]-1):
                for x in range(img2.size[0]-1):
                    a = pixdata[x+1,y][0]
                    b = pixdata[x-1,y][0]
                    c = pixdata[x,y+1][0]
                    d = pixdata[x,y-1][0]
                    if pixdata[x,y][0] <= 150 and (a+b<=250 or a+c<=250 or a+d<=250 or c+b<=250 or b+d<=250 or c+d<=250):
                        pass
                    else:
                        pixdata[x, y] = (255, 255, 255)
            # 去除最下面和最上边最右边的色块
            for x in range(img2.size[0]-8,img2.size[0]):
                for y in range(img2.size[1]):
                    pixdata[x, y] = (255, 255, 255)
            for x in range(img2.size[0]):
                pixdata[x,img2.size[1]-1] = (255, 255, 255)
                pixdata[x,img2.size[1]-2] = (255, 255, 255)
                pixdata[x,0] = (255, 255, 255)
                pixdata[x,1] = (255, 255, 255)
            img2.save('yanzhengma.png')

            # 识别验证码的文字并输入验证码
            img = driver.find_element_by_xpath('//div[@class="div_yzm"]/img[1]')
            text = pytesseract.image_to_string(Image.open('yanzhengma.png'))
            print(text)
            num = ''
            try:
                if text[1] == '-':
                    num = int(text[0]) - int(text[2])
            except (ValueError,IndexError):
                img.click()
                time.sleep(1)
                continue
            try:
                if text[1] == '+':
                    num = int(text[0]) + int(text[2])
            except (ValueError,IndexError):
                img.click()
                time.sleep(1)
                continue
            self.driver.find_element_by_xpath('//div[@class="div_yzm"]/input[1]').send_keys(str(num))
            time.sleep(1)
            self.driver.find_element_by_xpath('//input[@value="提交"]').click()
            time.sleep(1)
            self.driver.find_element_by_xpath('//div[@class="div_yzm"]/input[1]').clear()
            img.click()
            time.sleep(1)
            continue
        driver.quit()

if __name__ =="__main__":
    jf = Jiaofei()
    jf.read()
    driver = jf.write()
    jf.yanzheng(driver)

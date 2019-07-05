import xlrd
import os
import time,random
from appium import webdriver
from urllib3.exceptions import ProtocolError
from selenium.common.exceptions import NoSuchElementException

# 读取要添加的名单
def exists():
    if os.path.exists('biaoji.txt'):
        with open('biaoji.txt','r') as f:
            for index in f:
                return index
    return 0

def read(index):
    index = int(index.strip())
    with open('phones.txt','r',encoding='utf-8') as f:
        i = 0
        for line in f:
            i += 1
            if i <= index:
                continue
            yield line.split(',')

def add_friend(index):
    desired_caps = {'platformName': 'Android',
                    'deviceName': '587f66db3079',
                    'platformVersion': '4.4.4',#将要测试app的安装包放到自己电脑上执行安装或启动，如果不是从安装开始，则不是必填项，可以由下面红色的两句直接启动
                    'appPackage': 'com.tencent.mm', #红色部分如何获取下面讲解
                    'appActivity': 'com.tencent.mm.ui.LauncherUI',
                    'unicodeKeyboard': "True",    #使用unicode输入法
                    'resetKeyboard': "True",       #重置输入法到初始状态
                    'noReset': "True"
                    } 
    driver = webdriver.Remote('http://127.0.0.1:4723/wd/hub', desired_caps)
    while True:
        try:
            driver.find_element_by_xpath('//android.widget.RelativeLayout[@content-desc="更多功能按钮"]').click()
        except (NoSuchElementException,ProtocolError):
            time.sleep(5)
            continue
        break
    time.sleep(random.randint(1,3))
    # 点击添加朋友
    driver.find_elements_by_class_name('android.widget.TextView')[1].click()
    time.sleep(1)
    driver.find_elements_by_class_name('android.widget.TextView')[2].click()
    request = 0
    for i in read(index):
    # 点击添加微信好友
        phone = i[1].strip()
        time.sleep(random.randint(50,60))
        # 输入微信号码,此处做一个模拟人的输入
        driver.find_element_by_class_name('android.widget.EditText').send_keys(phone)
        # 点击搜索微信开始搜索
        time.sleep(random.randint(2,4))
        driver.find_element_by_class_name('android.widget.TextView').click()
        time.sleep(random.randint(7,12))
        try:
            if driver.find_elements_by_class_name('android.widget.TextView')[0].text == "该用户不存在":
                time.sleep(random.randint(2,4))
                driver.find_element_by_class_name('android.widget.EditText').clear()
                continue
            elif driver.find_elements_by_class_name('android.widget.TextView')[0].text == "被搜帐号状态异常，无法显示":
                time.sleep(random.randint(2,4))
                driver.find_element_by_class_name('android.widget.EditText').clear()
                continue
            elif driver.find_elements_by_class_name('android.widget.TextView')[0].text == "操作过于频繁，请稍后再试":
                time.sleep(2)
                driver.quit()
                return
            elif driver.find_elements_by_class_name('android.widget.TextView')[-3].text == "发消息":
                driver.find_elements_by_class_name('android.widget.ImageView')[-1].click()
                time.sleep(random.randint(1,2))
                driver.find_element_by_class_name('android.widget.EditText').clear()
                continue
            # 点击添加到通讯录
            driver.find_elements_by_class_name('android.widget.TextView')[-2].click()
        except (ProtocolError,IndexError,NoSuchElementException):
            continue
        # 点击发送请求消息
        time.sleep(random.randint(4,7))
        if request == 0:
            driver.find_elements_by_class_name('android.widget.EditText')[0].clear()
            driver.find_elements_by_class_name('android.widget.EditText')[0].send_keys('你好,我是专业协助政府无偿资助申报的公司')
        request = 1
        time.sleep(random.randint(3,6))
        driver.find_element_by_class_name('android.widget.Button').click()
        time.sleep(random.randint(3,5))
        # 点击返回键开始下一个添加
        driver.find_elements_by_class_name('android.widget.ImageView')[-1].click()
        # 清空搜索栏
        time.sleep(random.randint(2,4))
        try:
            driver.find_element_by_class_name('android.widget.EditText').clear()
        except NoSuchElementException:
            time.sleep(2)
            driver.quit()
            return
        with open('biaoji.txt','w',encoding='utf-8') as f:
            f.write(i[0].strip())

if __name__ == "__main__":
    index = exists()
    add_friend(index)

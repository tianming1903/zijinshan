from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException,TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
import time,sys,os
import xlrd,random,xlwt
from PIL import Image
import win32com

class Email(object):
    def __init__(self):
        self.biaozhi = input('需要传入附件吗(y/n):')
        self.email = input('输入邮箱号:')
        self.password = '19830913yX'
        self.driver = webdriver.Chrome()
        self.driver.get('https://qiye.aliyun.com/')

    # 获取发送记录
    def read_log(self):
        namelist = []
        if os.path.exists('./log.txt'):
            with open('./log.txt','r',encoding='utf-8') as f:
                for name in f.readlines():
                    if name.split('_')[0] not in namelist:
                        namelist.append(name)
            return len(namelist)+1
        return 1

     # 获取主题
    def read_text(self):
        filename = ''
        filelist = next(os.walk('./'))[2]
        l = [i for i in filelist if '.txt' in i]
        for i in l:
            if i != 'log.txt':
                filename = i
                break
        with open(filename,'r') as f:
            themes = f.readlines()
            while True:
                theme = random.choice(themes)
                if theme == '\n':
                    continue
                theme = theme.strip()
                break
            return theme

    # 获取word文件
    def copy_world(self):
        # 获取word文件
        words = []
        dirs = os.path.abspath('.')
        dirname = next(os.walk('./'))[1][-1]
        filelist = next(os.walk(dirname))[2]
        for i in filelist:
            if i.split('.')[1] == 'doc' or i.split('.')[1] == 'docx':
                words.append(i)
        wordname = random.choice(words)
        # 复制文档的所有内容
        app=win32com.client.Dispatch('Word.Application')
        # 打开word，经测试要是绝对路径
        doc=app.Documents.Open(dirs + '\\' + dirname + '\\' + wordname)
        # 复制word的所有内容
        doc.Content.Copy()
        # 关闭word
        doc.Close()

    # 获取图片
    def read_img(self):
        images = []
        # 获取文件家里面的所有文件
        dirname = next(os.walk('./'))[1][-1]
        filelist = next(os.walk(dirname))[2]
        # 选择里面的图片文件
        for i in filelist:
            if i.split('.')[1] == 'jpg' or i.split('.')[1] == 'png':
                images.append(i)
        # 删除掉 ‘发送文件’
        for i in images:
            if '发送' in i:
                os.remove(dirname + '/' + i)
                images.remove(i)
        filename = random.choice(images)
        img = Image.open(dirname + '/' + filename)
        # 生成一个倍率来改变图片的大小
        m = random.uniform(1.5,0.8)
        size = img.size
        width = int(size[0]*m)
        height = int(size[1]*m)
        pic = img.resize((width,height),Image.ANTIALIAS)
        g = random.choice(('.jpg','.png'))
        pic=pic.convert('RGB')
        pic.save(dirname + '/' + '发送' + g)
        return (dirname + '\\' + '发送' + g)
    
    # 获取附件
    def fujian(self):
        filelist = next(os.walk('./'))[2]
        for i in filelist:
            if '.doc' in i or '.docx' in i:
                return i

    # 读取要发送的邮箱
    def read_excel(self,num):
        emails = []
        files = next(os.walk('./'))[2]
        filename = ''
        for i in files:
            if ('xls' in i) or ('xlsx' in i):
                filename = i
                break
        book = xlrd.open_workbook(filename)    
        table = book.sheets()[0]
        
        # 更改邮箱位置
        name = table.col_values(0)[num:]
        faren = table.col_values(1)[num:]
        email = table.col_values(2)[num:]

        for x,y,z in zip(name,faren,email):
            for i in z:
                if ord(i) > 225:
                    break
            else:
                l = z.split(';')
                for e in l:
                    if len(e) > 5:
                        info = x + '_' + y + '<' + e.strip() + '>'
                        emails.append(info)
        return emails
        
    # 发送主题类型加工
    def send_them(self,thme,email):
        if ' ' not in thme:
            return thme
        name = email.split('<')[0].split('_')[1]
        theme = thme.replace("' '",name)
        return theme

    # 发送图片
    def send_images(self,image):
        dirs = os.path.abspath('.')
        time.sleep(random.randint(1,2))
        self.driver.find_element_by_xpath('//div[@class="editor_toolbar_line"]/div[@_id="imageupload"]').click()
        time.sleep(random.randint(3,5))
        self.driver.find_elements_by_xpath('//input[contains(@class,"uploader_input")]')[1].send_keys(dirs  + '\\' + image)
        time.sleep(random.randint(1,3))
        self.driver.find_element_by_xpath('//div[@_v="ok"]').click()
        # 点击发送，判断图片是否加载到文本处，切换到iframe中
        self.driver.switch_to_frame(self.driver.find_element_by_class_name('aym_editor_iframe'))
        while True:
            try:
                WebDriverWait(self.driver,5).until(EC.visibility_of_all_elements_located((By.XPATH,'//span/img')))
            except TimeoutException:
                continue
            break
        self.driver.switch_to_default_content()

    def send_word(self):
        # 选择文本输入
        self.driver.switch_to_frame(self.driver.find_element_by_class_name('aym_editor_iframe'))
        self.driver.find_element_by_xpath('//body').send_keys(Keys.CONTROL,'v')
        self.driver.switch_to_default_content()

    def login(self):
        # 切换到iframe页面去
        frame = self.driver.find_element_by_xpath('//iframe[@allowtransparency="true"]')
        self.driver.switch_to_frame(frame)
        frame = self.driver.find_element_by_xpath('//iframe[@id="ding-login-iframe"]')
        self.driver.switch_to_frame(frame)
        self.driver.find_element_by_xpath('//div[@class="login_section"]//input[@id="username"]').send_keys(self.email)
        time.sleep(2)
        self.driver.find_element_by_xpath('//div[@class="login_section"]//input[@id="password"]').send_keys(self.password)
        while True:
            try:
                password = self.driver.find_element_by_xpath('//div[@class="login_section"]//input[@id="password"]')
            except NoSuchElementException:
                pass
            else:
                if password.get_attribute('value') == '':
                    password.send_keys(self.password)
            try:
                WebDriverWait(self.driver,3).until(EC.visibility_of_all_elements_located((By.XPATH,'//div[@class="navbar_top"]/div[1]/b')))
            except TimeoutException:
                continue
            else:
                break

    def inputs(self,emails):
        index = 0
        l = []
        # 获取图片的路径
        for email in emails:
            if index >= 50:
                self.driver.quit()
                print('被退回邮件' + str(l[1] - l[0]) + '封')
                sys.exit('请更换邮箱')
            time.sleep(random.randint(4,7))
            
            # 获取未读邮件数量
            try:
                now_number = self.driver.find_element_by_xpath('//div[@class="treenode_children_wrap"]/div[1]/div/div').text.split('(')[1][:-1]
            except IndexError:
                now_number = '0'
            l.append(int(now_number))
            if len(l) == 3:
                if l[1] - l[0] >= 10:
                    self.driver.quit()
                    sys.exit('被退回的邮箱数量过多，请检查发送内容')
                else:
                    del l[1]
                    
            # 点击发送邮件
            self.driver.find_element_by_xpath('//div[@class="navbar_top"]/div[1]/b').click()
            time.sleep(random.randint(2,4))
            
            # 输入收件人,做一个判断是否有出现发送邮件的界面
            while True:
                try:
                    self.driver.find_element_by_xpath('//input[@class="entry_input evt" and @tabindex = "101"]').send_keys(email)
                except NoSuchElementException:
                    time.sleep(2)
                    self.driver.find_element_by_xpath('//div[@class="navbar_top"]/div[1]/b').click()
                    continue
                break

            # 更改主题并且输入主题
            them = self.read_text()
            theme = self.send_them(them,email)
            time.sleep(random.randint(1,3))
            self.driver.find_elements_by_xpath('//div[@class="textbox_inner"]/input')[1].send_keys(theme)
            
            # 输入附件
            if (self.biaozhi.strip()).lower() == 'y':
                time.sleep(random.randint(3,6))
                filename = self.fujian()
                dirs = os.path.abspath('./')
                self.driver.find_element_by_xpath('//input[@type="file"]').send_keys(dirs + '\\' + filename)
                while True:
                    if '成功' in self.driver.find_element_by_xpath('//span[@class="compose_att_stat_txt"]').text:
                        break
                    time.sleep(2)
                    continue

            # 选择内容形式(后面加上1即可对发送的内容作出改变，可以发送图片和文本)
            search = random.choice((0,))
            if search == 0:
                image = self.read_img()
                self.send_images(image)
            elif search == 1:
                time.sleep(random.randint(2,5))
                self.copy_world()
                self.send_word()
            
            # 点击发送
            time.sleep(random.randint(2,4))
            self.driver.find_element_by_xpath('//div[@class="compose_toolbar_abs"]/div[1]').click()
            
            # 判断是否发送成功
            time.sleep(random.randint(4,7))
            status = self.driver.find_elements_by_xpath('//div[@class="root_tabbar_items_wrap"]/div[@draggable="true"]')
            if len(status) == 2:
                self.driver.find_elements_by_xpath('//div[@_v="close"]')[1].click()
                time.sleep(random.randint(2,3))
                self.driver.find_element_by_xpath('//div[@_v="cancel"]').click()
            
            # 写入发送邮件的记录
            with open('log.txt','a',encoding='utf-8') as f:
                f.write(email)
                f.write('\n')
            index += 1
        os.remove('./log.txt')
        self.driver.quit()
        print('邮箱发送完毕，请更换')

if __name__ == "__main__":    
    ea = Email()
    num = ea.read_log()
    emails = ea.read_excel(num)
    ea.login()
    ea.inputs(emails)
    
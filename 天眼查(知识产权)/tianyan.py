# -*- coding: utf-8 -*-
import time,random
from Login import Login
from selenium.common.exceptions import NoSuchElementException,StaleElementReferenceException,TimeoutException
import xlwt,xlrd
import os,sys
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class TianyanSpider():
    def exists(self):
        #判断存储信息文件是否存在(info.txt)
        index = 0
        if os.path.exists('info.txt'):
            with open('info.txt','r',encoding='utf-8') as f:
                for x in f:
                    index = x.split(',')[0]
                return int(index)+1
        else:
            return 1

    def Identification(self):
        # 获取要爬取名单的文件名
        file = next(os.walk('./'))
        for filename in file[2]:
            if ('.xlsx' in filename) or ('.xls' in filename):
                return filename
        else:
            sys.exit('添加要爬取的excel名单')

    def read(self,number,filename):
        # 读取excel的公司并且获取名单
        '''
        excel文件的格式，只有名单这一列，而且必须从第二行开始
        '''
        data = xlrd.open_workbook(filename)
        sheet = data.sheets()[0]
        names = sheet.col_values(0)[number:]
        return names

    def login(self):
        # 登录
        object = Login().start()
        return object

    def search(self,object,names,number):
        # 按公司名进行搜索
        for name in names:
            name = name.strip()
            object.find_elements_by_xpath('//input[@type="search"]')[0].send_keys(name)
            time.sleep(random.randint(1,2))
            # 异常处理
            try:
                object.find_elements_by_xpath('//div[@class="name js-text"]')[0].click()
            except (NoSuchElementException,IndexError):
                object.find_elements_by_xpath('//input[@type="search"]')[0].clear()
                number += 1
                continue
            except StaleElementReferenceException:
                object.quit()
                sys.exit('程序被迫中止，请重新启动')
            self.fanhui(object,number)
            number += 1

    def fanhui(self,object,num):
        # 切换到新的页面进行数据的提取
        window = object.window_handles
        object.switch_to.window(window[1])
        self.extract(object,num)
        # 关闭新页面，切换到原页面输入下一个公司的名称进行查询
        object.close()
        object.switch_to.window(window[0])
        object.find_elements_by_xpath('//input[@type="search"]')[0].clear()

    def extract(self,object,num):
        # 进行数据的提取
        textlist = [str(num)]
        time.sleep(1)
        # 验证码的等待过程
        while True:
            try:
                name = object.find_element_by_xpath('//div[@class="header"]/h1').text
            except NoSuchElementException:
                try:
                    object.find_element_by_xpath('//div[@class="container -body"]')
                except NoSuchElementException:
                    time.sleep(5)
                    continue
                return
            textlist.append(name)
            break

        # 商标信息的抓取
        shangbiao_list = []
        try:
            shangbiao = object.find_element_by_xpath('//div[contains(@tyc-event-ch,"shangbiao")]')
        except NoSuchElementException:
            pass
        else:
            while True:
                infos = shangbiao.find_elements_by_xpath('.//tbody/tr')
                for i in infos:
                    info = ''
                    try:
                        l = i.find_elements_by_xpath('./td/span')
                        for x in l:
                            info += x.text + ','
                    except StaleElementReferenceException:
                        continue
                    shangbiao_list.append(info.strip(','))
                try:
                    next_button = shangbiao.find_element_by_xpath('.//a[@class="num -next"]')
                except NoSuchElementException:
                    break
                object.execute_script('arguments[0].click();',next_button)
                time.sleep(1)

        # 专利信息的抓取
        zhuanli_list = []
        try:
            zhuanli = object.find_element_by_xpath('//div[contains(@tyc-event-ch,"zhuanli")]')
        except NoSuchElementException:
            pass
        else:
            while True:
                infos = zhuanli.find_elements_by_xpath('.//tbody/tr')
                for i in infos:
                    info = ''
                    try:
                        l = i.find_elements_by_xpath('./td/span')
                        for x in l:
                            info += x.text + ','
                    except StaleElementReferenceException:
                        continue
                    zhuanli_list.append(info.strip(','))
                try:
                    next_button = zhuanli.find_element_by_xpath('.//a[@class="num -next"]')
                except NoSuchElementException:
                    break
                object.execute_script('arguments[0].click();',next_button)
                time.sleep(1)

        # 软件著作权的抓取
        zhuzuoquan_list = []
        try:
            zhuzuoquan = object.find_element_by_xpath('//div[contains(@tyc-event-ch,"zhuzuoquan")]')
        except NoSuchElementException:
            pass
        else:
            while True:
                infos = zhuzuoquan.find_elements_by_xpath('.//tbody/tr')
                for i in infos:
                    info = ''
                    try:
                        l = i.find_elements_by_xpath('./td/span')[:-1]
                        for x in l:
                            info += x.text + ','
                    except StaleElementReferenceException:
                        continue
                    zhuzuoquan_list.append(info.strip(','))
                try:
                    next_button = zhuzuoquan.find_element_by_xpath('.//a[@class="num -next"]')
                except NoSuchElementException:
                    break
                object.execute_script('arguments[0].click();',next_button)
                time.sleep(1)
        
        # 作品信息的抓取
        zuoping_list = []
        try:
            zuoping = object.find_element_by_xpath('//div[contains(@tyc-event-ch,"zuopinzhuzhuoquan")]')
        except NoSuchElementException:
            pass
        else:
            while True:
                infos = zuoping.find_elements_by_xpath('.//tbody/tr')
                for i in infos:
                    info = ''
                    try:
                        l = i.find_elements_by_xpath('./td/span')
                        for x in l:
                            info += x.text + ','
                    except StaleElementReferenceException:
                        continue
                    zuoping_list.append(info.strip(','))
                try:
                    next_button = zuoping.find_element_by_xpath('.//a[@class="num -next"]')
                except NoSuchElementException:
                    break
                object.execute_script('arguments[0].click();',next_button)
                time.sleep(1)

        # 对网站备案的爬取
        icp_list = []
        try:
            icp = object.find_element_by_xpath('//div[contains(@tyc-event-ch,"Icp")]')
        except NoSuchElementException:
            pass
        else:
            while True:
                infos = object.find_elements_by_xpath('//div[contains(@tyc-event-ch,"Icp")]//tbody/tr')

                for i in infos:
                    info = ''
                    try:
                        l = i.find_elements_by_xpath('./td/span')
                        for x in l:
                            info += x.text + ','
                    except StaleElementReferenceException:
                        continue
                    yuming = i.find_element_by_xpath('./td[5]').text
                    wangzhi = i.find_element_by_xpath('./td//a').text
                    info += wangzhi + ',' + yuming
                    icp_list.append(info)
                try:
                    next_button = icp.find_element_by_xpath('.//a[@class="num -next"]')
                except NoSuchElementException:
                    break
                object.execute_script('arguments[0].click();',next_button)
                time.sleep(1)
        with open('info.txt','a',encoding='utf-8') as f:
            f.write(','.join(textlist))
            f.write('\n')
        if shangbiao_list == [] and zhuanli_list == [] and zhuzuoquan_list == [] and zuoping_list == [] and icp_list == []:
            return
        self.write(textlist,shangbiao_list,zhuanli_list,zhuzuoquan_list,zuoping_list,icp_list)

    def write(self,textlist,shangbiao_list,zhuanli_list,zhuzuoquan_list,zuoping_list,icp_list):
        # 创建一个excel表
        now = time.strftime('%m%d%H%M%S', time.localtime(time.time()))
        book = xlwt.Workbook()
        sheet = book.add_sheet('my_book')
         # 创建表格样式
        style = xlwt.XFStyle()
        alignment = xlwt.Alignment()
        alignment.horz = xlwt.Alignment.HORZ_CENTER  #水平方向居中
        alignment.vert = xlwt.Alignment.VERT_CENTER  #垂直方向居中
        style.alignment = alignment
        for i in range(7):
            if i == 2:
                sheet.col(i).width = 256 * 60
            else:
                sheet.col(i).width = 256 * 20
        status = 1
        sheet.write_merge(0,0,0,6,textlist[1],style=style)

        #把信息写入到excel
        if shangbiao_list != []:
            sheet.write_merge(status+1,status + len(shangbiao_list),0,0,'商标信息',style=style)
            sheet.write(status, 1, '申请日期',style=style)
            sheet.write(status, 2, '商标名称',style=style)
            sheet.write(status, 3, '注册号',style=style)
            sheet.write(status, 4, '国际分类',style=style)
            sheet.write(status, 5, '商标状态',style=style)
            x = status + 1
            for info in shangbiao_list:
                text = info.split(',')
                for i in range(len(text)):
                    sheet.write(x,i+1,text[i])
                x += 1 
            status += len(shangbiao_list) + 2

        if zhuanli_list != []:
            sheet.write_merge(status+1,status + len(zhuanli_list),0,0,'专利信息',style=style)
            sheet.write(status, 1, '申请公布日',style=style)
            sheet.write(status, 2, '专利名称',style=style)
            sheet.write(status, 3, '申请号',style=style)
            sheet.write(status, 4, '申请公布号',style=style)
            sheet.write(status, 5, '专利类型',style=style)
            x = status + 1
            for info in zhuanli_list:
                text = info.split(',')
                for i in range(len(text)):
                    sheet.write(x,i+1,text[i])
                x += 1 
            status += len(zhuanli_list) + 2

        if zhuzuoquan_list != []:
            sheet.write_merge(status+1,status + len(zhuzuoquan_list),0,0,'软件著作权信息',style=style)
            sheet.write(status, 1, '批准日期',style=style)
            sheet.write(status, 2, '软件全称',style=style)
            sheet.write(status, 3, '软件简称',style=style)
            sheet.write(status, 4, '登记号',style=style)
            sheet.write(status, 5, '分类号',style=style)
            sheet.write(status, 6, '版本号',style=style)
            x = status + 1
            for info in zhuzuoquan_list:
                text = info.split(',')
                for i in range(len(text)):
                    sheet.write(x,i+1,text[i])
                x += 1 
            status += len(zhuzuoquan_list) + 2

        if zuoping_list != []:
            sheet.write_merge(status+1,status + len(zuoping_list),0,0,'作品著作权信息',style=style)
            sheet.write(status, 1, '作品名称',style=style)
            sheet.write(status, 2, '登记号',style=style)
            sheet.write(status, 3, '作品类别',style=style)
            sheet.write(status, 4, '创作完成日期',style=style)
            sheet.write(status, 5, '登记日期',style=style)
            sheet.write(status, 6, '首次发表日期',style=style)
            x = status + 1
            for info in zuoping_list:
                text = info.split(',')
                for i in range(len(text)):
                    sheet.write(x,i+1,text[i])
                x += 1 
            status += len(zuoping_list) + 2
        
        if icp_list != []:
            sheet.write_merge(status+1,status + len(icp_list),0,0,'网站备案',style=style)
            sheet.write(status, 1, '审核日期',style=style)
            sheet.write(status, 2, '网站名称',style=style)
            sheet.write(status, 3, '网站许可',style=style)
            sheet.write(status, 4, '网站地址',style=style)
            sheet.write(status, 5, '域名',style=style)
            x = status + 1
            for info in icp_list:
                text = info.split(',')
                for i in range(len(text)):
                    sheet.write(x,i+1,text[i])
                x += 1 
        if os.path.exists('爬取好的'):
            pass
        else:
            os.mkdir('爬取好的')
        book.save('爬取好的/' + textlist[1] + now + '.xls')
    
    def main(self):
        index = self.exists()
        name = self.Identification()
        names = self.read(index,name)
        object = self.login()
        self.search(object,names,index)
        os.remove('info.txt')

if __name__ == "__main__":
    spider = TianyanSpider()
    spider.main()
    print('名单爬取完毕')
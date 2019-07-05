from selenium import webdriver
import os,sys
import time 
import xlrd,xlwt
from xlrd.biffh import XLRDError
from selenium.common.exceptions import NoSuchElementException

def request():
    # 修改文件的保存地址
    # 创建一个文件夹
    dirs = os.path.abspath('./') + '\\' + 'mingdan'
    os.mkdir(dirs)
    options = webdriver.ChromeOptions()
    prefs = {'profile.default_content_settings.popups':0,
            'download.default_directory':dirs}
    options.add_experimental_option('prefs',prefs)
    # 发送请求获取链接
    driver = webdriver.Chrome(chrome_options= options)
    driver.get('http://hb.beian.miit.gov.cn/state/outPortal/moreLatestMessage.action')

    # 加载更多的页面
    try:
        time.sleep(2)
        driver.find_element_by_xpath('//option[5]').click()
    except NoSuchElementException:
        driver.quit()
        delete()
        sys.exit('504错误请重新运行')

    # 获取时间
    with open('time.txt','r') as f:
        now_time = f.readlines()[0].strip()
    release_times = driver.find_elements_by_xpath('//tbody[2]/tr/td[3]')
    num = 1
    for x in release_times:
        if now_time == x.text.strip():
            break
        num += 1
    # 要存储的时间，供下次比对
    r_time = release_times[0].text
    if num == 1:
        delete()
        sys.exit('名单已经是最新，没有必要抓取')
    l = []
    # 收集名单额链接，进行下载
    for i in range(1,num):
        name = driver.find_element_by_xpath('//tbody/tr[@id='+ str(i) +']/td[2]').text
        if ('(' or ')') in name:
            link = driver.find_element_by_xpath('//tbody/tr[@id='+ str(i) +']/td[4]/a').get_attribute('href')
            l.append(link)

    # 对每个下载文件地址进行请求访问
    for i in l:
        driver.get(i)
        time.sleep(2)
        try:
            driver.find_element_by_xpath('//span/a').click()
        except NoSuchElementException:
            time.sleep(4)
            continue
        time.sleep(4)
    time.sleep(10)
    driver.quit()
    return r_time

# 读取所有的excel文件保存公司的名称
def inputs():
    print('开始读取下载好的excel表')
    files = next(os.walk('./mingdan'))[2]
    for i in files:
        try:
            book = xlrd.open_workbook('./mingdan/' + i)
        except XLRDError:
            continue
        for table in book.sheets():
            names = table.col_values(1)[2:]
            with open('namelist.txt','a',encoding='utf-8') as f:
                for name in names:
                    if str(name).strip()[-4:] == '有限公司':
                        f.write(str(name))
                        f.write('\n')

# 读取名单文件进行去重
def quchong():
    print('开始去重')
    namelist = []
    with open('namelist.txt','r',encoding='utf-8') as f:
        for name in f:
            if name.strip() not in namelist:
                 namelist.append(name.strip())
    with open('name.txt','a',encoding='utf-8') as f:
        for name in namelist:
            f.write(name)
            f.write('\n')
    os.remove('namelist.txt')
    return namelist

#分批次读取name.txt保存为excel格式
def readinfo(filename = 'name.txt'):
    text = []
    with open (filename,'r',encoding='utf-8') as f:
        for name in f:
            text.append(name)
            if len(text) ==2000:
                yield text
                text = []
        else:
            if text == []:
                return
            yield text

#把信息写入到excel
def write(r_time):
    print('开始把名单写入excel表中')
    for text in readinfo():
        time.sleep(1)
        now = time.strftime('%m%d%H%M%S', time.localtime(time.time()))
        book = xlwt.Workbook()
        sheet = book.add_sheet('my_book')
        sheet.write(0, 0, '公司名称')
        x = 1
        for infos in text:
            sheet.write(x,0,infos)
            x += 1
        if os.path.exists('./爬取好的'):
            pass        
        else:
            os.mkdir('./爬取好的')
        book.save('./爬取好的/' + './ICP备案' + now + '.xls')
    with open('time.txt','w',encoding='utf-8') as f:
        f.write(r_time)
    os.remove('name.txt')

def delete():
    print('开始删除不必要文件')
    files = next(os.walk('./mingdan'))
    for i in files[2]:
        os.remove('./mingdan/' + i)
    os.rmdir('./mingdan')

if __name__ == "__main__":
    t1 = time.time()
    r_time = request()
    inputs()
    namelist = quchong()
    write(r_time)
    delete()
    t2 = time.time()
    print('名单抓取完成')
    print('成功抓取数据' + str(len(namelist)) + '条')
    print('耗时' + str(t2-t1) + '秒')

from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC 
from selenium.webdriver.common.by import By
import time,random
from PIL import Image
import pytesseract
from io import BytesIO
from selenium.common.exceptions import NoSuchElementException,NoAlertPresentException,TimeoutException,ElementNotVisibleException
import os,sys
import xlrd

class Zhishi(object):
	def __init__(self):
		self.driver = webdriver.Chrome()
		self.driver.get('http://cpservice.sipo.gov.cn/index.jsp')
		self.driver.maximize_window()
	
	def read_text(self):
		files = next(os.walk('./'))[2]
		num = 0
		for f in files:
			if '.txt' in f:
				with open(f,'r',encoding='utf-8') as f:
					num = len(f.readlines())	
		return num

	def read_excel(self,num):
		files = next(os.walk('./'))[2]
		for f in files:
			if '.xls' in f or '.xlsx' in f:
				break
		book = xlrd.open_workbook(f)
		table = book.sheets()[0]
		length = table.col_values(0)
		for index in range(num+1,len(length)):
			info = table.row_values(index)
			self.login(info)
			self.liucheng()
			self.write(info)
		self.driver.quit()

	def login(self,info):
		# 进行登录
		try:
			WebDriverWait(self.driver,10).until(EC.visibility_of_element_located((By.XPATH,'//div[@class="mlogin"]//li')))
		except TimeoutException:
			sys.exit('网络较差，结束填写任务')
		# 填写账户和密码
		self.driver.find_element_by_xpath('//div[@class="mlogin"]//li[1]/input').send_keys(info[0])
		time.sleep(random.randint(1,2))
		self.driver.find_element_by_xpath('//div[@class="mlogin"]//li[2]/input').send_keys(info[1])
		# 输入验证码
		while True:
			# 获取全屏
			screenshot = self.driver.get_screenshot_as_file('screenshot.png')
			screenshot = Image.open('screenshot.png')

			# 获取验证码的位置和大小
			img = self.driver.find_element_by_xpath('//img[@class="wrpyzm"]')
			location = img.location
			size = img.size
			top,bottom,left,right = location['y'], location['y'] + size['height'], location['x'], location['x'] + size['width']
			
			# 截取验证码
			yanzhengma = screenshot.crop((left, top, right, bottom))
			yanzhengma.save('yanzhengma.png')
			
			# 对照片进行底色处理并保存
			image = Image.open('./yanzhengma.png')
			img2 = image.convert('RGBA')
			pixdata = img2.load()
			for y in range(img2.size[1]):
				for x in range(img2.size[0]):
					if pixdata[x,y][0]>= 110:
						pixdata[x, y] = (255, 255, 255)
			img2.save('yanzhengma.png')
			
			# 识别验证码的文字并输入验证码
			text = pytesseract.image_to_string(Image.open('./yanzhengma.png'))
			self.driver.find_element_by_xpath('//div[@class="mlogin"]//li[3]/input').send_keys(text)
			time.sleep(1)
			self.driver.find_element_by_xpath('//input[@class="loginbtn"]').click()
			time.sleep(2)
			# 如果没有通过则刷新验证码重新验证
			try:
				alert = self.driver.switch_to_alert()
			except NoAlertPresentException:
				break
			else:
				alert.accept()
				self.driver.find_element_by_xpath('//input[@class="wrpinput3"]').clear()
				continue

	def liucheng(self):
		# 统一协议
		try:
			WebDriverWait(self.driver,5).until(EC.element_to_be_clickable((By.XPATH,'//input[@id="agreeid"]')))
		except TimeoutError:
			sys.exit('网络太差，结束填写')
		self.driver.find_element_by_xpath('//input[@id="agreeid"]').click()
		time.sleep(1)
		self.driver.find_element_by_xpath('//input[@id="next"]').click()

		# 点击业务办理
		time.sleep(2)
		self.driver.find_element_by_xpath('//div[@class="hmenu"]//li[2]//img').click()
		time.sleep(2)
		self.driver.find_element_by_xpath('//h1[@class="menu1"][4]').click()

		# 切换到iframe中
		self.driver.switch_to_frame(self.driver.find_element_by_xpath('//iframe[@id="rightFrame"]'))
		try:
			WebDriverWait(self.driver,10).until(EC.element_to_be_clickable((By.XPATH,'//h1/input[@class="inconbutb"]')))
		except TimeoutException:
			sys.exit('网络太差，结束填写')
		self.driver.find_element_by_xpath('//h1/input[@class="inconbutb"]').click()
		
		# 点击协议并且同意
		time.sleep(2)
		self.driver.find_element_by_xpath('//input[@id="tongyi"]').click()
		time.sleep(1)
		self.driver.find_element_by_xpath('//input[@id="goBtn"]').click()
	
	# 填写企业类型的费减
	def write(self,info):
		# 点击企业类型
		while True:
			try:
				self.driver.find_element_by_xpath('//span[@class="xzwcontarl"]//input[@value="2"]').click()
			except NoSuchElementException:
				time.sleep(2)
				continue
			break
		time.sleep(3)
		# 勾选2019
		self.driver.find_element_by_xpath('//input[@value="2019"]').click()
		time.sleep(1)
		# 填写企业名称
		self.driver.find_element_by_xpath('//input[@name="beianrmc"]').send_keys(info[2])
		time.sleep(1)
		# 填写信用代码
		nums = self.driver.find_elements_by_xpath('//input[@name="qiYeJGDm_in"]')
		for i,y in zip(nums,info[0]):
			i.send_keys(y)
		time.sleep(1)
		# 填写企业人数
		self.driver.find_element_by_xpath('//input[@name="qiyecyrs"]').send_keys(str(info[3]).split('.')[0])
		time.sleep(1)
		# 填写资产总额
		self.driver.find_element_by_xpath('//input[@name="zichanze"]').send_keys(str(info[4]))
		time.sleep(1)
		# 填写年度应纳税金额
		self.driver.find_element_by_xpath('//input[@name="ndynssde"]').send_keys(str(info[5]))
		time.sleep(1)
		
		# 如何根据地址解决正确填写省、市、区、县
		separator = ['省','市','区','县']
		index_list = []
		string = info[6]
		x = 0
		lentgh = 3 if '省' in info[6] else 2
		for i in info[6]:
			if x == lentgh:
				break
			if i in separator:
				index_list.append(string.split(i)[0])
				string = string.split(i)[1]
				x += 1

		# 填写省或者直辖市
		zhucedsfdm = self.driver.find_element_by_xpath('//select[@name="zhucedsfdm"]')
		zhucedsfdm.click()
		for i in zhucedsfdm.find_elements_by_xpath('./option'):
			if index_list[0] in i.text:
				i.click()
		time.sleep(0.5)

		# 填写市
		zhucedcsdm = self.driver.find_element_by_xpath('//select[@name="zhucedcsdm"]')
		zhucedcsdm.click()
		for i in zhucedcsdm.find_elements_by_xpath('./option'):
			if index_list[1] in i.text:
				i.click()
		time.sleep(0.5)

		# 填写县，做一个判断是否有
		try:
			zhucedxjdm = self.driver.find_element_by_xpath('//select[@name="zhucedxjdm"]')
			zhucedxjdm.click()
		except ElementNotVisibleException:
			pass
		else:
			for i in zhucedxjdm.find_elements_by_xpath('./option'):
				if index_list[2] in i.text:
					i.click()
		time.sleep(1)

		# 输入地址和电话号码以及联系人
		self.driver.find_element_by_xpath('//input[@name="xiangxidz"]').send_keys(info[6])
		time.sleep(1)
		self.driver.find_element_by_xpath('//input[@name="lianxirmc"]').send_keys(info[7])
		time.sleep(1)
		self.driver.find_element_by_xpath('//input[@name="lianxirdh"]').send_keys(str(info[8]).split('.')[0])
		time.sleep(1)

		# 对联系人地址做出切割
		index_list = []
		string = info[9]
		x = 0
		lentgh = 3 if '省' in info[9] else 2
		for i in info[9]:
			if x == lentgh:
				break
			if i in separator:
				index_list.append(string.split(i)[0])
				string = string.split(i)[1]
				x += 1

		# 输入联系人地址
		lianxirsfdm = self.driver.find_element_by_xpath('//select[@name="lianxirsfdm"]')
		lianxirsfdm.click()
		for i in lianxirsfdm.find_elements_by_xpath('./option'):
			if index_list[0] in i.text:
				i.click()
		time.sleep(0.5)
		lianxircsdm = self.driver.find_element_by_xpath('//select[@name="lianxircsdm"]')
		lianxircsdm.click()
		for i in lianxircsdm.find_elements_by_xpath('./option'):
			if  index_list[1] in i.text:
				i.click()
		time.sleep(0.5)
		try:
			lianxirxjdm = self.driver.find_element_by_xpath('//select[@name="lianxirxjdm"]')
			lianxirxjdm.click()
		except ElementNotVisibleException:
			pass
		else:
			for i in lianxirxjdm.find_elements_by_xpath('./option'):
				if  index_list[2] in i.text:
					i.click()
		time.sleep(1)

		# 输入联系人地址
		self.driver.find_element_by_xpath('//input[@name="lianxirdz"]').send_keys(info[9])
		time.sleep(1)

		# 点击预览
		self.driver.find_element_by_xpath('//input[@id="yulanButton"]').click()
		time.sleep(2)

		# 点击提交和确定
		self.driver.find_element_by_xpath('//input[@id="submitButton"]').click()
		time.sleep(1)
		self.driver.find_element_by_xpath('//button[contains(@class,"ui-button")]').click()
		# 把成功的名单写入到文件做好备录
		with open('chenggong.txt','a',encoding="utf-8") as f:
			f.write(info[2])

		# 点击退出，先返回到住iframe中
		self.driver.switch_to_default_content()
		time.sleep(1)
		self.driver.find_element_by_xpath('//div[@class="flright"]/a').click()
		time.sleep(1)
		self.driver.find_element_by_xpath('//div[@class="ui-dialog-buttonset"]/button[1]/span').click()
		
if __name__ == "__main__":
	zs = Zhishi()
	num = zs.read_text()
	zs.read_excel(num)
	print('所有的名单费减已经完成')


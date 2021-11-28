import os
import time
import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, UnexpectedAlertPresentException


class BatteryWeb(object):

    def __init__(self, url, title, browser='Firefox'):
        self.url = url
        self.title = title
        self.browser = browser
        self.driver = self.open_browser()
        self.open_url()

    def open_browser(self):
        try:
            #chromedriver.exe、geckodriver.exe等放置在C:\Program Files\Python310\Scripts
            if self.browser == 'Chrome':
                return webdriver.Chrome()
            elif self.browser == 'Ie':
                return webdriver.Ie()
            elif self.browser == 'Firefox':
                options = webdriver.FirefoxOptions()
                options.add_argument('-headless')
                return webdriver.Firefox(firefox_options=options)
            elif self.browser == 'Edge':
                return webdriver.Edge()
        except Exception as e:
            print("启动浏览器异常：", e)
            return None

    def get_element(self, value=None, by='id'):
        try:
            if by == 'id':
                return self.driver.find_element_by_id(value)
            elif by == 'name':
                return self.driver.find_element_by_name(value)
            elif by == 'className':
                return self.driver.find_element_by_class_name(value)
            elif by == 'xpath':
                return self.driver.find_element_by_xpath(value)
            else:
                return self.driver.find_element_by_css(value)
        except Exception as e:
            print("find_element错误信息：", e)
            return False

    def get_screenshot(self):
        picture_time = time.strftime("%Y%m%d-%H%M%S", time.localtime(time.time()))
        #路径：当前目录上一级目录下results/images/时间.png
        picture_path = os.path.abspath(os.path.join(os.getcwd(), "..", 'Data', 'images', picture_time+'.png'))
        print('异常信息已保存：{}'.format(picture_path))
        return self.driver.get_screenshot_as_file(picture_path)

    def open_url(self):
        #self.driver.maximize_window()
        self.driver.get(self.url)
        self.driver.implicitly_wait(10)
        if EC.title_contains(self.title)(self.driver):
            print('打开网址成功')
            return self.driver
        else:
            self.get_screenshot()
            return None

    def web_login(self, username, password):
        self.get_element('username', 'name').clear()
        self.get_element('username', 'name').send_keys(username)
        self.get_element('password', 'name').send_keys(password)
        self.get_element('//*[@id="app"]/div/div/div/form/div[4]/div/button', 'xpath').click()
        self.driver.implicitly_wait(3)

        try:
            so =WebDriverWait(self.driver,5,0.5).until(EC.visibility_of_element_located((By.XPATH,'/html/body/div[2]/p')))
            #WebDriverWait(self.driver, 5, 0.2).until(EC.alert_is_present())
            #alert = self.driver.switch_to.alert
            print('Alert提示信息：{}'.format(so.text))
            #alert.accept()  # 去除浏览器警告
            self.get_screenshot()
        except TimeoutException:
            if self.get_element('//*[@id="app"]/div/div/div/div/ul/div/a/li/span', 'xpath').text == '首页':
                return True
            else:
                return self.get_screenshot()



    def get_validBatRow1(self):
        '''
        获取查询电池信息有效果的行数
        return: 电池信息总行数，有效果电池信息行数，GRPS在线状态电池个数
        '''
        #解决新版本没有数据时等待时间超过1min的问题
        start_time = time.time()
        element1 = '/html/body/div/div/div/section/div/div/div[2]/div/div[3]/div/span'
        try:
            self.driver.implicitly_wait(5)
            WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located((By.XPATH, element1)))
            #print(element_none)
            print(time.time()-start_time)
            #element_none = self.driver.find_element_by_xpath(element1)
            return [0, 0, 0]
        except TimeoutException:
        #except NoSuchElementException:
            if time.time() - start_time > 2:
                return [1, 1, 1]


    def get_validBatRow(self):
        '''
        获取查询电池信息有效果的行数
        return: 电池信息总行数，有效果电池信息行数，GRPS在线状态电池个数
        '''
        row = 0  # 有效电池信息行数
        sum = 0  # 电池信息和
        num = 0  # GPRS在线电池个数
        updated = ''
        validBatRow = [sum, row, num]

        # 解决新版本没有数据时等待时间超过1min的问题
        menu_table = self.get_element('//*[@id="app"]/div/div/section/div/div/div[2]/div/div[3]/table', 'xpath')
        try:
            self.driver.implicitly_wait(2)
            #n = len(menu_table.find_elements_by_tag_name('tr'))  # 获取电池状态table行数
            n = len(menu_table.find_elements(by=By.TAG_NAME, value='tr'))
            if n == 0:
                return [0, 0, 0]
            elif n == 1:
                batStatus = self.get_batStatus(1, 1)
                # 判断数据有效性：PACK包含6020或6442、更新时间非空
                if (batStatus[0][0:6] == 'BR6442' or batStatus[0][0:6] == 'BR6020') and (batStatus[6] != ''):
                    if batStatus[2] == '在线':
                        return [1, 1, 1]
                    else:
                        return [1, 1, 0]
            else:
                for i in range(1, n + 1):
                    batStatus = self.get_batStatus(n, i)
                    if (batStatus[0][0:6] == 'BR6442' or batStatus[0][0:6] == 'BR6020') and (batStatus[6] > updated):
                        updated = batStatus[6]
                        row = i
                        sum = n
                        if batStatus[2] == '在线':
                            num = num + 1
                    return [sum, row, num]
        except TimeoutException:
            return [0, 0, 0]


    def query_batStatus(self, value, type='3', by='devid'):
        '''
        查询电池状态: 
        :param value:  devid value, imei value, imsi value

        :param by: devid、imei、imsi
        '''        
        self.driver.implicitly_wait(60)
        element1 = '//*[@id="app"]/div/div/div/div/ul/div/li[2]/div/i'
        if not self.driver.find_element_by_xpath('//*[@id="app"]/div/div/div/div/ul/div/li[2]/ul/a[2]/li/span').is_displayed():
            self.get_element(element1, 'xpath').click()
            self.driver.implicitly_wait(5)

        WebDriverWait(self.driver, 60).until(EC.element_to_be_clickable((By.XPATH, element1)))
        WebDriverWait(self.driver, 60).until(EC.visibility_of_element_located((By.XPATH, element1)))

        element2 = self.driver.find_element_by_xpath('//*[@id="app"]/div/div/div/div/ul/div/li[2]/ul/a[2]/li/span')
        webdriver.ActionChains(self.driver).move_to_element(element2).click(element2).perform()
        while True:
            if self.driver.find_element_by_xpath('//*[@id="app"]/div/ul/div[2]/span/span[3]/span[1]/span').text == '电池列表':
                break
            else:
                webdriver.ActionChains(self.driver).move_to_element(element2).click(element2).perform()
                self.driver.implicitly_wait(5)
        self.get_element('tab-{}'.format(type), 'id').click()
        self.get_element('//*[@id="app"]/div/div/section/div/div/form/div[2]/div/div/button[1]', 'xpath').click()
        self.get_element('//*[@id="app"]/div/div/section/div/div/form/div[2]/div/div/div[4]/div/div/input', 'xpath').send_keys(value)
        path = '//*[@id="app"]/div/div/section/div/div/form/div[2]/div/div/div[4]/div/div/input'
        while True:
            if self.get_element(path, 'xpath').get_attribute('value') == value:
                break
            else:
                self.get_element(path, 'xpath').clear()
                time.sleep(1)
                self.get_element('//*[@id="app"]/div/div/section/div/div/form/div[2]/div/div/button[2]/span',
                                 'xpath').click()
                time.sleep(1)
                self.get_element(path, 'xpath').send_keys(value)
        self.get_element('//*[@id="app"]/div/div/section/div/div/form/div[2]/div/div/button[2]/span','xpath').click()
        time.sleep(2)

    def get_batStatus(self, sum, row):
        '''
        获取电池状态信息: DEVID、IMEI、IMSI、GPRS状态、BMS状态、当前电池、当前场景、当前厂商、最后上报时间
        :param sum: 获取电池状态行数和
        :param row: 获取电池状态行数
        '''

        batStatus = []
        for i in [3, 4, 5, 6, 7, 8, 9, 12]:
            if sum == 1:
                path = '//*[@id="app"]/div/div/section/div/div/div[2]/div/div[3]/table/tbody/tr/td[{}]/div'.format(i)
            else:
                path = '//*[@id="app"]/div/div/section/div/div/div[2]/div/div[3]/table/tbody/tr[{}]/td[{}]/div'.format(row,i)
            value = self.get_element(path, 'xpath').text
            batStatus.append(value)
        if batStatus[7] == '':
            batStatus.append("NG")
        else:
            datetime_object = datetime.datetime.strptime(batStatus[7], '%Y-%m-%d %H:%M:%S')
            time_stamp = datetime.datetime.now()
            time = (time_stamp - datetime_object).total_seconds() - 12600  #超过3.5h未更新，判定异常
            if time > 0:
                batStatus.append("NG")
            else:
                batStatus.append("OK")
        return batStatus

    def get_batWebInfo(self, sum, row):
        '''
        电池详细信息查询：
        :param sum: 电池状态行数和
        :param row: 电池信息行数
        return: devid,imsi,imei,bms_pcb,bms_ver,pcb_ver,gprs_ver,voltage,soc,soh,capacity
        '''

        if sum == 1:
            path = '//*[@id="app"]/div/div/section/div/div/div[2]/div/div[3]/table/tbody/tr/td[13]/div/div[1]/button[1]/span'
        else:
            path = '//*[@id="app"]/div/div/section/div/div/div[2]/div/div[3]/table/tbody/tr[{}]/td[13]/div/div[1]/button[1]/span'.format(row)
        self.get_element(path, 'xpath').click()
        self.driver.implicitly_wait(15)
        time.sleep(2)
        i = 2
        j = 1
        batteryInfo = []
        list = [(1,6,0),(3,2,0),(3,4,0),(3,6,1),(2,8,1),(4,4,1),(6,2,0)]
        for i,j,k in list:
            if k == 0:
                path = '//*[@id="app"]/div/div/section/div/div/div[2]/div/div/div[{}]/div[{}]'.format(i,j)
            else:            
                path = '//*[@id="app"]/div/div/section/div/div/div[2]/div/div/div[{}]/div[{}]/span[{}]'.format(i,j,k)
            value = self.get_element(path, 'xpath').text
            batteryInfo.append(value)
        return batteryInfo

    def get_batBoxTableNum(self):
        '''
        查询电柜 柜中电池数量:返回表的数量
        '''
        self.driver.implicitly_wait(5)
        if not self.driver.find_element_by_xpath('//*[@id="app"]/div/div/div/div/ul/div/li[1]/ul/a[4]/li/span').is_displayed():
            self.get_element('//*[@id="app"]/div/div/div/div/ul/div/li[1]/div/i', 'xpath').click()
            self.driver.implicitly_wait(10)
        so = WebDriverWait(self.driver,30).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="app"]/div/div/div/div/ul/div/li[1]/ul/a[4]/li/span')))
        so.click()
        self.driver.implicitly_wait(5)
        table_min = self.driver.find_element_by_xpath(
            '//*[@id="app"]/div/div/section/div/div/div[4]/div/span[3]/div/input').get_attribute('min')
        table_max = self.driver.find_element_by_xpath(
            '//*[@id="app"]/div/div/section/div/div/div[4]/div/span[3]/div/input').get_attribute('max')
        return [table_min,table_max]

    def get_batBoxTableRow(self,sn):
        '''
        查询电柜 柜中电池每页（表）中电池个数
        '''
        self.get_element('//*[@id="app"]/div/div/section/div/div/div[4]/div/span[3]/div/input', 'xpath').clear()
        time.sleep(1)
        self.get_element('//*[@id="app"]/div/div/section/div/div/div[4]/div/span[3]/div/input', 'xpath').clear()
        #time.sleep(3)
        self.get_element('//*[@id="app"]/div/div/section/div/div/div[4]/div/span[3]/div/input', 'xpath').send_keys(sn)
        time.sleep(1)
        self.get_element('//*[@id="app"]/div/div/section/div/div/div[4]/div/span[3]/div/input', 'xpath').send_keys(
            Keys.ENTER)
        self.driver.implicitly_wait(10)
        time.sleep(1)
        menu_table = self.get_element('//*[@id="app"]/div/div/section/div/div/div[3]/div/div[3]/table', 'xpath')
        row = len(menu_table.find_elements_by_tag_name('tr'))
        return row

    def get_batBoxTableInfo(self, row):
        '''
        查询电柜 柜中电池Info
        '''
        batBoxTableInfo = []
        for k in range(1,7):
            value = self.get_element('//*[@id="app"]/div/div/section/div/div/div[3]/div/div[3]/table/tbody/tr[{}]/td[{}]/div'.format(row,k),'xpath').text
            batBoxTableInfo.append(value)
        return batBoxTableInfo

    def web_logout(self):
        self.driver.quit()
        print('关闭浏览器成功')

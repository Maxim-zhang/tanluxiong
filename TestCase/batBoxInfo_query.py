import time
from selenium import webdriver
from Tanluxiong.Lib import battery_web
from Tanluxiong.Lib import battery_excel
from Tanluxiong.Conf import battery_config


class batWebInfoQuery(object):

    def __init__(self, filename, sheetname):

        cf = battery_config.BatConfig()
        self.username = cf.get_batConfig('tanluxiong', 'tlx_username')
        self.password = cf.get_batConfig('tanluxiong', 'tlx_password')
        self.url = cf.get_batConfig('tanluxiong', 'tlx_server')
        self.title = cf.get_batConfig('tanluxiong', 'tlx_title')
        #方式一：通过main配置
        self.filename = filename
        self.sheetname = sheetname
        #方式二：通过配置文件读取
        #self.filename = cf.get_batConfig('battery_excel', 'filename')
        #self.sheetname = cf.get_batConfig('battery_excel', 'sheetname')
        self.ws = battery_excel.BatteryExcel(self.filename, self.sheetname)
        self.driver = battery_web.BatteryWeb(self.url, self.title)
        

    def Info_query(self):
        # 登录GPS管理系统
        if self.driver.web_login(self.username, self.password):
            print('用户:{} 登录成功'.format(self.username))
        else:
            print('用户:{} 登录失败'.format(self.username))
        tableNum = self.driver.get_batBoxTableNum()

        for i in range (int(tableNum[0]), int(tableNum[1])+1):
            j = self.driver.get_batBoxTableRow(i)
            for row in range(1,j):
                batBoxInfo = self.driver.get_batBoxTableInfo(row)
                self.ws.write_batBoxTableInfo((i-1)*10+row+1, batBoxInfo)
                print('第{}页 {}行查询成功:{}'.format(i, row, batBoxInfo))
                

    def web_close(self):
    
        self.driver.web_logout()


if __name__ == '__main__':
    #设置EXCEL文件名称、表名称，文件默认保存在Tanluxiong/Data目录
    
    rsp = batWebInfoQuery(filename='Battery11071.xlsx', sheetname='BOX')
    rsp.Info_query()
    rsp.web_close()
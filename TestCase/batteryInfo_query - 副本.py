import time
from selenium import webdriver
from Tanluxiong.Lib import battery_web
from Tanluxiong.Lib import battery_excel
from Tanluxiong.Conf import battery_config


class batWebInfoQuery(object):

    def __init__(self, filename, sheetname):

        cf = battery_config.BatConfig()
        #self.username = cf.get_batConfig('tanluxiong', 'tlx_username')
        #self.password = cf.get_batConfig('tanluxiong', 'tlx_password')
        self.username = cf.get_batConfig('tanluxiong', 'bat_username')
        self.password = cf.get_batConfig('tanluxiong', 'bat_password')
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

    def Info_query(self, start=2, end=1000, by='devid'):
        # 登录GPS管理系统
        if self.driver.web_login(self.username, self.password):
            print('用户:{} 登录成功'.format(self.username))
        else:
            print('用户:{} 登录失败'.format(self.username))
        #默认值设置表格的最大行数
        if end == 1000:
            end == self.ws.get_batTableRow()

        for n in range(start, end+1):
            value = ''
            remarks = ''
            batStatus = []
            batWebInfo = []
            batInfo = self.ws.read_batInfo(n)
            if by == 'devid':
                value = batInfo[0]
            elif by == 'imei':
                value = batInfo[2]
            else:
                value = batInfo[3]
            if value is None:
                pass
            else:
                self.driver.query_batStatus(value, by)
                validBatRow = self.driver.get_validBatRow()
                if validBatRow[1] > 0:
                    batStatus = self.driver.get_batStatus(sum=validBatRow[0], row=validBatRow[1])
                    batWebInfo = self.driver.get_batWebInfo(sum=validBatRow[0], row=validBatRow[1])
                    #EXCEL表、状态页面、详细页面三处查询到到DEVID、IMEI、IMSI不一致时添加备注
                    for (a,b,c) in  [(batInfo[0],batStatus[0],batWebInfo[0]),(batInfo[2],batStatus[1],batWebInfo[2]),(batInfo[3],batStatus[2],batWebInfo[1])]:
                        if (a == b and b == c) or (a =='' and b == c):
                            continue
                        else:
                            remarks = 'DEVID/IMEI/IMSI不一致'
                    #通信IMEI、IMSI查询时，在线个数超过1个时添加备注
                    if validBatRow[2] > 1:
                        remarks = 'GPRS在线电池数超过1个'
                    batStatus.append(remarks)
                #查询数据为空时处理：DEVID、BMS_PCB、IMEI、IMSI保持不变，其它内容置空，备注暂无数据
                else:
                    for i in [0, 2, 3]:
                        batStatus.append(batInfo[i])
                    for j in range(4, 11):
                        batStatus.append('')
                    batStatus.append('暂无数据')
                    batStatus[9]= "NG"
                    for k in range(0,11):
                        batWebInfo.append('')

                if batStatus and batWebInfo:
                    print('第 {} 行查询成功：{}'.format(n,value))
                    self.ws.write_batStatus(n, batStatus)
                    self.ws.write_batInfo(n, batWebInfo)
                    self.ws.save_batTable()
                else:
                    print('第 {} 行查询失败,请确认！！！：{}'.format(n, value))

    def web_close(self):
    
        self.driver.web_logout()

if __name__ == '__main__':
    #设置EXCEL文件名称、表名称，文件默认保存在Tanluxiong/Data目录
    
    file = 'Battery11071.xlsx'
    by = 'devid'
    cfg1 = [file, 'S50', 2, 51, by]
    cfg2 = [file, 'S60', 2, 61, by]
    cfg3 = [file, 'L36', 2, 37, by]
    cfg4 = [file, 'L46', 2, 47, by]
    cfg5 = [file, 'L142', 2, 143, by]
    cfg6 = [file, 'L192', 2, 193, by]
    cfg_all = [cfg6]

    for cfg in cfg_all:
        rsp = batWebInfoQuery(filename=cfg[0], sheetname=cfg[1])
        #设置查询开始行数、结束行数、查询方式：devid、imei、imsi
        rsp.Info_query(start=cfg[2], end=cfg[3], by=cfg[4])
        rsp.web_close()

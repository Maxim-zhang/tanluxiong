import os
#import re
import openpyxl
from openpyxl import load_workbook
from Tanluxiong.Conf.battery_config import BatConfig
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE


class BatteryExcel(object):

    def __init__(self, filename, sheetname=None):
        """
        初始化方法
        :param filename: Excel文件名，默认存放在Data目录
        :param sheetname: 表单名
        """
        self.bat_path = os.path.abspath(os.path.abspath('..') + '/Data/' + filename)
        if not os.path.exists(self.bat_path):
            raise FileNotFoundError('请确保电池文件存在:{}'.format(self.bat_path))
        self.sheetname = sheetname
        self.wb = load_workbook(self.bat_path)
        if self.sheetname is None:
            self.ws = self.wb.active # active默认读取第一个表单
        else:
            self.ws = self.wb[self.sheetname] # 读取指定表单

    def read_batInfo(self, row):
        '''
        Battery info读取：devid、imsi、imei、场景
        :param row: Excel表行数
        '''
        battery_info = []
        for i in range (1, 5):
            info = self.ws.cell(row, column=i).value
            battery_info.append(info)
        return battery_info

    def write_batInfo(self, row, batInfo):
        '''
        Battery info写入,bms_ver,pcb_ver,gprs_ver,voltage,soc,soh,capacity
        :param row: Excel表行数
        :param batInfo：battery info
        '''
        for i in range(11, 15):
            self.ws.cell(row=row, column=i, value=batInfo[i-8]) #写入bms_ver,gprs_ver,soc,soh,capacity
        self.ws.cell(row=row, column=2, value=batInfo[2])
        self.ws.cell(row=row, column=3, value=batInfo[1])


    def write_batStatus(self, row, batStatus):
        '''
        batteryStatus写入：IMEI、IMSI、GPRS状态、BMS状态、当前电流、当前场景、当前厂商、最后上报时间、remarks
        :param row: Excel表行数
        '''
        for i in range(6, 11):
            self.ws.cell(row=row, column=i, value=batStatus[i-3])  #写入IMEI、IMSI、GPRS状态、当前电流、当前场景、当前厂商、最后上报时间
        #self.ws.cell(row=row, column=1, value=batStatus[0])  # 写入DEVID
        self.ws.cell(row=row, column=17, value=batStatus[9]) #写入remarks
        self.ws.cell(row=row, column=18, value=batStatus[8])  # 写入4G状态
        self.ws.cell(row=row, column=4, value=batStatus[10])  # 写入场景
        

    def get_batTableRow(self):
        return self.ws.max_row #获取最后一行行号
        

    def write_batBoxTableInfo(self, Row, value):
        for i in range (1,7):
            text = ILLEGAL_CHARACTERS_RE.sub(r'', value[i-1]) #写入之前，先将非法字符用空格代替
            self.ws.cell(row=Row, column=i, value=text)
        self.wb.save(self.bat_path)  


    def save_batTable(self):
        self.wb.save(self.bat_path)
        self.wb.close()






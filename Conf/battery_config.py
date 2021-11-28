from configparser import ConfigParser
import os


class BatConfig(object):

    def __init__(self):
        self.bat_config = ConfigParser()
        config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'battery_config.ini')
        if not os.path.exists(config_path):
            raise FileNotFoundError('请确保配置文件存在:{}'.format(config_path))
        self.bat_config.read(config_path, encoding='utf-8')

    def get_batConfig(self, section, option):
        """
        配置文件读取
        :param section:
        :param option:
        """
        return self.bat_config.get(section, option)
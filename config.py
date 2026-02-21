# config.py
import configparser
import os
import sys

# 1. 确定 settings.ini 的绝对路径
# 【终极防翻车技巧】这里加上了应对未来打包 exe 的特殊判断。
# 如果被打包成了 exe，sys.frozen 就会生效，它会去 exe 所在的文件夹找 ini；
# 如果是在 PyCharm 里运行，它就去当前代码所在的文件夹找 ini。
if getattr(sys, 'frozen', False):
    base_path = os.path.dirname(sys.executable)
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

ini_path = os.path.join(base_path, 'settings.ini')

# 2. 检查配置文件在不在，不在直接报错罢工
if not os.path.exists(ini_path):
    raise FileNotFoundError(f"❌ 找不到配置文件！请确保 {ini_path} 文件存在。")

# 3. 召唤解析器，读取 ini 文件
# encoding='utf-8' 极其关键，防止中文注释或中文路径变成乱码报错
config = configparser.ConfigParser()
config.read(ini_path, encoding='utf-8')

# 4. 把读出来的值，赋给原本的全局变量
ERP_URL = config.get('Network', 'ERP_URL')
USERNAME = config.get('Account', 'USERNAME')
PASSWORD = config.get('Account', 'PASSWORD')

SOURCE_EXCEL_PATH = config.get('Files', 'SOURCE_EXCEL_PATH')
OUTPUT_EXCEL_NAME = config.get('Files', 'OUTPUT_EXCEL_NAME')

# 这个通常不变，就不用独立出去了，依然留在代码里
COLUMN_NAME_CODE = "项目编号"
"""
ValkyrieEngine 全局配置解析模块
功能：负责在程序启动时定位并读取外部的 settings.ini 配置文件。
该模块内置了环境自适应逻辑，确保无论是在源码环境还是打包后的独立可执行文件 (.exe) 环境中，
均能准确读取同级目录下的配置文件，并将配置参数映射为全局变量，供其他业务模块调用。
"""

import configparser
import os
import sys

# =========================================================
# 1. 配置文件路径解析与环境自适应
# =========================================================
# 判定当前程序的运行环境状态：
# getattr(sys, 'frozen', False) 用于检测程序是否被 PyInstaller 打包。
# 若为 True（处于 exe 运行环境），基准路径设定为 .exe 文件所在的绝对物理目录；
# 若为 False（处于源码运行环境），基准路径设定为当前 config.py 脚本所在的目录。
if getattr(sys, 'frozen', False):
    base_path = os.path.dirname(sys.executable)
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

# 拼接得出配置文件的最终绝对路径
ini_path = os.path.join(base_path, 'settings.ini')


# =========================================================
# 2. 配置文件校验与加载
# =========================================================
# 前置安全校验：若配置文件缺失，则直接抛出文件未找到异常，阻止引擎启动。
if not os.path.exists(ini_path):
    raise FileNotFoundError(f"[系统异常] 缺失核心配置文件，请确保 {ini_path} 文件存在。")

# 初始化配置解析器实例
config = configparser.ConfigParser()

# 读取配置文件数据
# 必须显式指定 encoding='utf-8'，以防止 Windows 系统默认编码读取中文注释或中文路径时引发解码崩溃。
config.read(ini_path, encoding='utf-8')


# =========================================================
# 3. 全局配置变量映射 (动态配置项)
# =========================================================
# 网络环境配置
ERP_URL = config.get('Network', 'ERP_URL')

# 账户鉴权配置
USERNAME = config.get('Account', 'USERNAME')
PASSWORD = config.get('Account', 'PASSWORD')

# 数据 I/O 路径配置 - 批量映射各个功能的专属路径
# 功能1
F1_INPUT = config.get('Files', 'F1_INPUT', fallback='')
F1_OUTPUT = config.get('Files', 'F1_OUTPUT', fallback='')
# 功能2
F2_INPUT = config.get('Files', 'F2_INPUT', fallback='')
F2_OUTPUT = config.get('Files', 'F2_OUTPUT', fallback='')
# 功能3
F3_INPUT = config.get('Files', 'F3_INPUT', fallback='')
F3_OUTPUT = config.get('Files', 'F3_OUTPUT', fallback='')
# 功能4
F4_INPUT = config.get('Files', 'F4_INPUT', fallback='')
F4_OUTPUT = config.get('Files', 'F4_OUTPUT', fallback='')
# 功能5
F5_INPUT = config.get('Files', 'F5_INPUT', fallback='')
F5_OUTPUT = config.get('Files', 'F5_OUTPUT', fallback='')
# 功能6
F6_INPUT = config.get('Files', 'F6_INPUT', fallback='')
F6_OUTPUT = config.get('Files', 'F6_OUTPUT', fallback='')


# =========================================================
# 4. 内部静态常量 (非动态配置项)
# =========================================================
# 指定目标 Excel 表格中用于检索的唯一标识列的表头名称。
# 由于此设定属于业务强相关且变动频率极低的基础规则，因此保留在代码内部，不暴露给外部用户修改。
COLUMN_NAME_CODE = "项目编号"
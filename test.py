import platform
import win32api, win32con
import Test_Scripts.EasyShell_Lib as EasyshellLib
import Library.CommonLib as CommonLib
import os
import time
import traceback
import pysnooper
from PIL import ImageGrab, ImageDraw, ImageFont
import subprocess
import chardet
import uiautomation
import random


def logoff():
    os.system('echo test_logon_pass >test_rdp_result.txt')
    ftp = CommonLib.FTPUtils('15.83.251.201', 'sh\\cheng.balance', 'password.321')
    ftp.change_dir('/Function/Automation/log/rdp')
    ftp.upload_file('test_rdp_result.txt', 'test_rdp_result.txt')
    ftp.close()
    os.system('shutdown -l')


a = os.path.basename('/Function/Automation/log/rdp')
print(a)
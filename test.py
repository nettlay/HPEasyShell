import time
from win32com.client import DispatchEx
from Test_Scripts import easyshell
import os


# def logoff():
#     os.system('echo test_logon_pass >test_rdp_result.txt')
#     ftp = easyshell.CommonLib.FTPUtils('15.83.251.201', 'sh\\cheng.balance', 'password.321')
#     ftp.change_dir('/Function/Automation/log/rdp')
#     ftp.upload_file('test_rdp_result.txt', 'test_rdp_result.txt')
#     ftp.close()
#     os.system('shutdown -l')

import uiautomation

print(uiautomation.TextControl(AutomationId='LabelTime', searchFromControl=uiautomation.WindowControl(Name="HP Easy Shell")).IsOffScreen)
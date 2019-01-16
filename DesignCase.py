# import os
# # import ruamel.yaml as yaml
from EasyShell_Lib import *
# from easyshell import *
import time
import uiautomation
import win32con, win32api

import email

# try:
#     subprocess.Popen('cmd')
#     subprocess.Popen('regedit')
#     sid = QATools.GetSid()
#     root = winreg.HKEY_USERS
#     path = r"{}\\Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings".format(sid)
#     with open('c:\\svc\\error.log','a') as f:
#         f.write(path)
#     key = winreg.OpenKeyEx(root, path, access=winreg.KEY_ALL_ACCESS | winreg.KEY_WOW64_64KEY)
#     winreg.SetValueEx(key, 'ProxyServer', 0, winreg.REG_SZ, '15.85.199.199:8080')
#     winreg.SetValueEx(key, 'ProxyOverride', 0, winreg.REG_SZ, '*.sh.dto;15.83.*.*')
#     winreg.SetValueEx(key, 'ProxyEnable', 0, winreg.REG_DWORD, 1)
# except Exception as e:
#     with open('c:\\svc\\error.log','a') as f:
#         f.write('\n{}'.format(e))
print(UserKiosk_Dict['MACAddr'].Name)
print(QATools.getNetInfo()['MAC'].replace('-', ":"))
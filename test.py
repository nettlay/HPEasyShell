# import shutil
# import time
# from win32com.client import DispatchEx
# from Test_Scripts import easyshell
import os, sys
import shutil

import openpyxl
import yaml
# import ruamel.yaml as yaml
import re


# def logoff():
#     os.system('echo test_logon_pass >test_rdp_result.txt')
#     ftp = easyshell.easyshell.CommonLib.FTPUtils('15.83.251.201', 'sh\\cheng.balance', 'password.321')
#     ftp.change_dir('/Function/Automation/log/rdp')
#     ftp.upload_file('test_rdp_result.txt', 'test_rdp_result.txt')
#     ftp.close()
#     os.system('shutdown -l')
# test = "Copy right 2012-2019 HP develepoment"
# print(re.findall(r'.*-\d{4} (.*)', test))


# def list_item(file):
#     cls = 'Base'
#     rs = {cls: []}
#     with open(file, 'r', encoding='utf-8') as f:
#         lines = f.readlines()
#     i = 0
#     for line in lines:
#         i += 1
#         # print(line[0:20])
#         if line.strip().startswith("class "):
#             cls_name = line.replace('class ', '').strip()[:-1]
#             cls = cls_name
#             rs[cls] = []
#             print(i)
#         if line.strip().startswith('def'):
#             func_name = line.replace('def ', '').strip()[:-1]
#             rs[cls].append(func_name)
#             print(i)
#     return rs

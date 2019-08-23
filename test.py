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

os.popen('HPEasyshellSettings.reg')
time.sleep(1)
for i in range(5):
    print(EasyshellLib.getElement('REG_EDIT_WARNNING').Exists(0, 0))
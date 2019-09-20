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

print(os.path.exists(r'c:\windows\sysnative\autologcfg.exe'))
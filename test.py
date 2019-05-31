from Test_Scripts.EasyShell_Lib import *

import uiautomation
# t = uiautomation.TextControl(Name='Hostname').Click()
import time
time.sleep(3)
print(uiautomation.GetFocusedControl())

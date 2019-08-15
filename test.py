# import time
# import os
import uiautomation
# import platform
#
#
# print(platform.system(),platform.version().split(".")[0])
# test = ['1', '12', '33']
# print(time.ctime())
print(uiautomation.TextControl(AutomationId='currentTitle', Name='Settings').Exists())
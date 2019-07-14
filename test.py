import Test_Scripts.EasyShell_Lib as EasyshellLib
import Library.CommonLib as CommonLib

# if EasyshellLib.getElement('MAIN_WINDOW').Exists(1, 1):
#     print('111')
#     EasyshellLib.getElement('MAIN_WINDOW').Close()
import uiautomation
import time
t = EasyshellLib.getElement('SoundSetting').Exists()
print(t)
EasyshellLib.getElement('WirelessSetting').Close()
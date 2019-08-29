import platform
import re

import win32api, win32con
import Test_Scripts.EasyShell_Lib as EasyshellLib
import Library.CommonLib as CommonLib
import os
import time
import traceback
import pysnooper
from PIL import ImageGrab, ImageDraw, ImageFont


def ClearContent(length=50):
    for temp in range(length):
        CommonLib.SendKey(CommonLib.Keys.VK_DELETE, 0.01)


class EasyShellTest:
    """
    profile is a test configuration name, it is a dict for app options, load from easyshell_testdata.yaml
    """

    def __init__(self):
        # ------- test root folder - ------------------
        self.path = EasyshellLib.file_path
        self.section_name = 'createApp'
        #  ---------------------------------
        self.log_path = os.path.join(self.path, 'Test_Report')
        self.misc = os.path.join(self.path, 'Misc')
        self.data = os.path.join(self.path, 'Test_Data')
        self.casepath = os.path.join(self.path, 'Test_Suite')
        self.testing = os.path.join(self.path, 'Testing')
        self.testset = os.path.join(self.path, 'testset.xlsx')
        self.sections = CommonLib.YmlUtils(os.path.join(self.data, "easyshell_testdata.yaml")).get_item()
        self.appPath = self.sections['appPath']['easyShellPath']
        self.debug = os.path.join(self.log_path, 'debug.log')

    def launch(self):
        if not EasyshellLib.getElement('MAIN_WINDOW').Exists(1, 1):
            for app_path in self.appPath:
                if os.path.exists(app_path):
                    EasyshellLib.CommonUtils.LaunchAppFromFile(app_path)
                    break
                else:
                    continue
            for i in range(20):
                if not EasyshellLib.getElement('MAIN_WINDOW').Exists():
                    continue
                else:
                    print('get window')
                    time.sleep(10)
                    break

    def resetEasyshell(self):
        easyshell_data = CommonLib.YmlUtils(os.path.join(self.data, 'standard.yaml'))
        app = easyshell_data.get_sub_item('app')
        rdp = easyshell_data.get_sub_item('rdp')
        site = easyshell_data.get_sub_item('site')
        store = easyshell_data.get_sub_item('store')
        settings = easyshell_data.get_sub_item('settings')
        reg = CommonLib.Reg_Utils()
        # clear settings
        if reg.isKeyExist(r'software\hp\hp easy shell\connections\rdp'):
            reg.clear_subkeys(r'software\hp\hp easy shell\connections\rdp')
        if reg.isKeyExist(r'software\hp\hp easy shell\connections\VMware'):
            reg.clear_subkeys(r'software\hp\hp easy shell\connections\VMware')
        if reg.isKeyExist(r'software\hp\hp easy shell\connections\CitrixICA'):
            reg.clear_subkeys(r'software\hp\hp easy shell\connections\CitrixICA')
        if reg.isKeyExist(r'software\hp\hp easy shell\Apps'):
            reg.clear_subkeys(r'software\hp\hp easy shell\Apps')
        if reg.isKeyExist(r'software\hp\hp easy shell\Sites'):
            reg.clear_subkeys(r'software\hp\hp easy shell\Sites')
        if reg.isKeyExist(r'software\hp\hp easy shell\StoreFront'):
            reg.clear_subkeys(r'software\hp\hp easy shell\StoreFront')
        # create temp settings
        general_key = reg.isKeyExist(r'software\hp\hp easy shell')
        if general_key:
            reg.create_value(general_key, 'KioskMode', 0, 'True')
            reg.create_value(general_key, 'KioskModeAdmin', 0, 'False')
            reg.create_value(general_key, 'VirtualKeyboardStyle', 0, '0')
            reg.create_value(general_key, 'DelayStart', 0, '0')
            reg.close(general_key)
        # -----------App__________________________
        reg.create_key(r'software\hp\hp easy shell\Apps\app0')
        app_key = reg.isKeyExist(r'software\hp\hp easy shell\Apps\app0')
        for name, value in app.items():
            reg.create_value(app_key, name, 0, value)
        reg.close(app_key)
        # ---------RDP --------------------------
        reg.create_key(r'software\HP\HP Easy Shell\connections\RDP\test_rdp')
        rdp_key = reg.isKeyExist(r'software\HP\HP Easy Shell\connections\RDP\test_rdp')
        for name, value in rdp.items():
            reg.create_value(rdp_key, name, 0, value)
        reg.close(rdp_key)
        reg.create_key(r'software\HP\HP Easy Shell\groups\test_rdp')
        # -----------web sites -------------------------
        reg.create_key(r'software\HP\HP Easy Shell\sites\site0')
        site_key = reg.isKeyExist(r'software\HP\HP Easy Shell\sites\site0')
        for name, value in site.items():
            reg.create_value(site_key, name, 0, value)
        reg.close(site_key)
        # -----------store -------------------------
        reg.create_key(r'software\HP\HP Easy Shell\StoreFront\test_store')
        store_key = reg.isKeyExist(r'software\HP\HP Easy Shell\StoreFront\test_store')
        for name, value in store.items():
            reg.create_value(store_key, name, 0, value)
        reg.close(store_key)
        # -----------settings -------------------------
        setting_key = reg.isKeyExist(r'software\HP\HP Easy Shell\UI')
        for name, value in settings.items():
            reg.create_value(setting_key, name, 0, value)
        reg.close(setting_key)

    def setKioskAdmin(self):
        reg = CommonLib.Reg_Utils()
        key = reg.isKeyExist(r'SOFTWARE\HP\HP Easy Shell')
        if key:
            reg.create_value(key, 'KioskModeAdmin', 0, 'True')
            reg.create_value(key, 'KioskMode', 0, 'True')
            reg.close(key)

    def clearKioskAdmin(self):
        reg = CommonLib.Reg_Utils()
        key = reg.isKeyExist(r'HKEY_LOCAL_MACHINE\SOFTWARE\HP\HP Easy Shell')
        if key:
            reg.create_value(key, 'KioskModeAdmin', 0, 'False')

    def create(self, profile):
        pass

    def edit(self, newprofile, oldprofile):
        pass

    def check(self, profile):
        pass

    # ------------------------------ Utils -------------------------------------------
    def Logfile(self, rs):
        if not os.path.exists(self.log_path):
            os.mkdir(self.log_path)
        EasyshellLib.TxtUtils(os.path.join(self.log_path, "easyShellLog.txt"), 'a').set_msg(
            "[{}]:{}\n".format(time.ctime(), rs))

    def capture(self, filename, txt=""):
        if not os.path.exists(self.log_path):
            os.mkdir(self.log_path)
        if not os.path.exists(os.path.join(self.log_path, 'screenshot')):
            os.mkdir(os.path.join(self.log_path, 'screenshot'))
        name = filename
        im = ImageGrab.grab()
        draw = ImageDraw.Draw(im)
        fnt = ImageFont.truetype(r'C:\Windows\Fonts\Tahoma.TTF', 24)
        draw.text((5, 5), txt, fill='red', font=fnt)
        for i in range(20):
            if os.path.exists("{}\\screenshot\\{}.jpg".format(self.log_path, name)):
                name = filename + str(i)
            else:
                break
        im.save("{}\\screenshot\\{}.jpg".format(self.log_path, name))

    # @pysnooper.snoop(os.path.join(__file__, 'Test_report\\debug.log'))
    def utils(self, profile='', op='exist', item='normal'):
        """
        :param profile:  test profile, [test1,test2,,standardApp...]
        :param op: test option [exist | notexist | shown | edit | delete |launch |default]
        :param item: specific for connection, if item=connection, connection button element=textcontrol.getparent,
                else element=textcontrol.getparent.getparent
        :return: Bool
        """
        time.sleep(3)
        test = self.sections[self.section_name][profile]
        name = test["Name"]
        if op.upper() == 'NOTEXIST':
            if CommonLib.TextControl(Name=name).Exists(0, 0):
                self.Logfile('Check {}-{} Not Exist Fail'.format(profile, name))
                self.capture(profile, 'Check {}-{} Not Exist Fail'.format(profile, name))
                return False
            else:
                self.Logfile('Check {}-{} Not Exist PASS'.format(profile, name))
                return True
        if CommonLib.TextControl(Name=name).Exists(0, 0):
            txt = CommonLib.TextControl(Name=name)
        else:
            print("{}-{} Not Exist".format(profile, name))
            return False
        if 'CONN' in item.upper():
            appControl = txt.GetParentControl()
        else:
            appControl = txt.GetParentControl().GetParentControl()
        launch = txt  # type=txtcontrol
        if appControl.ButtonControl(AutomationId='editButton').Exists(1, 1):
            edit = appControl.ButtonControl(AutomationId='editButton')
        else:
            edit = appControl.ButtonControl(AutomationId='EditButton')
        if appControl.ButtonControl(AutomationId='deleteButton').Exists(1, 1):
            delete = appControl.ButtonControl(AutomationId='deleteButton')
        else:
            delete = appControl.ButtonControl(AutomationId='DeleteButton')
        if op.upper() == 'LAUNCH':
            launch.Click()
            return True
        elif op.upper() == 'EDIT':
            edit.Click()
            return True
        elif op.upper() == 'SHOWN':
            if txt.IsOffScreen:
                return False
            else:
                return True
        elif op.upper() == 'DEFAULT':
            print('home')
            appControl.ButtonControl(AutoamtionId='homeButton').Click()
        elif op.upper() == 'DELETE':
            try:
                delete.Click()
                EasyshellLib.getElement('DeleteYes').Click()
                EasyshellLib.getElement('APPLY').Click()
                return True
            except:
                self.Logfile("[FAIL]:App {} Delete\nErrors:\n{}\n".format(name, traceback.format_exc()))
                self.capture(profile, "[FAIL]:App {} Delete\nErrors:\n{}\n".format(name, traceback.format_exc()))
                return False
        else:
            return True


class UserInterfacSettings(EasyShellTest):
    def __init__(self):
        EasyShellTest.__init__(self)
        self.section_name = 'userInterface'

    def edit(self, profile):
        """
        :param profile: one of test parameters' combination
        """
        flag = True  # record the function's status
        test = self.sections[self.section_name][profile]
        self.launch()
        self.Logfile('---------Begin To Edit User Interface settings----------')
        EasyshellLib.getElement('Settings').Click()
        EasyshellLib.getElement('KioskMode').Enable()
        for item in test:
            name = item.split(":")[0].strip()  # setting name
            status = item.split(":")[1].strip()  # setting status on/off
            if status.upper() == 'ON':
                try:
                    EasyshellLib.getElement(name).Enable()
                except:
                    flag = False
                    self.Logfile('[Fail]Button {} Enable\n{}'.format(name, traceback.format_exc()))
                    self.capture(profile, 'Button {} Enable\n{}'.format(name, traceback.format_exc()))
            elif status.upper() == 'OFF':
                try:
                    EasyshellLib.getElement(name).Disable()
                except:
                    flag = False
                    self.Logfile('[Fail]Button {} Disable\n{}'.format(name, traceback.format_exc()))
                    self.capture(profile, '[Fail]Button {} Disable\n{}'.format(name, traceback.format_exc()))
            else:
                flag = False
                self.Logfile('[Fail]Button {} Status in test data is not Correct!'.format(name))
                self.capture(profile, '[Fail]Button {} Status in test data is not Correct!'.format(name))
        EasyshellLib.getElement('APPLY').Click()
        EasyshellLib.getElement('Exit').Click()
        self.Logfile('[PASS] Modify user interface settings')
        return flag

    def check(self, profile):
        flag = True
        # ---Below test use yml, list type can be test in order --------
        content = self.sections
        test = content['userInterface'][profile]
        # 以下为部分设置的总开关sw = switch
        if EasyshellLib.getElement('MAIN_WINDOW').Exists():
            EasyshellLib.getElement('MAIN_WINDOW').SetFocus()
        swTitles = True
        swBrowser = True
        swPower = True
        self.Logfile('---------Begin To Check User Interface settings----------')
        for item in test:
            """
            1. 没有测试总开关关闭但是子开关开启时的状态
            2. Wifi 只测试链接窗口是否弹出，没有测试wifi连接功能
            """
            name = item.split(":")[0].strip()
            status = item.split(":")[1].strip()
            try:
                if status == "ON":
                    # ///////判断主按钮是否关闭，如果OFF,子按钮不再检查
                    if not swTitles:
                        if name in ['DisplayApp', 'DisplayConnections', 'DisplayStoreFront', 'DisplayWebsites']:
                            continue

                    if not swBrowser:
                        if name in ['DisplayAddress', 'DisplayNavigation', 'DisplayHome']:
                            continue
                    if not swPower:
                        if name in ['AllowLock', 'AllowLogoff', 'AllowRestart', 'AllowShutDown']:
                            continue
                    # ///////////////////////////////////
                    if name == 'DisplayTitle':
                        if not EasyshellLib.getElement('UserTitles').IsOffScreen:
                            self.Logfile("[PASS]: {} is shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is not shown".format(name))
                            self.capture(profile, "[Fail]: {} is not shown".format(name))
                    if name == 'DisplayBrowser':
                        if not EasyshellLib.getElement('UserBrowser').IsOffScreen:
                            self.Logfile("[PASS]: {} is shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is not shown".format(name))
                            self.capture(profile, "[Fail]: {} is not shown".format(name))
                    if name == 'DisplayAdmin':
                        if not EasyshellLib.getElement('UserAdmin').IsOffScreen:
                            self.Logfile("[PASS]: {} is shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is not shown".format(name))
                            self.capture(profile, "[Fail]: {} is not shown".format(name))
                    if name == 'DisplayPower':
                        if not EasyshellLib.getElement('UserPower').IsOffScreen:
                            self.Logfile("[PASS]: {} is shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is not shown".format(name))
                            self.capture(profile, "[Fail]: {} is not shown".format(name))
                    # ////////// item for Titles //////////////////////////////////////
                    if name == 'DisplayApp':
                        EasyshellLib.getElement('UserTitles').Click()
                        time.sleep(1)
                        if not EasyshellLib.getElement('UserApp').IsOffScreen:
                            self.Logfile("[PASS]: {} is shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is not shown".format(name))
                            self.capture(profile, "[Fail]: {} is not shown".format(name))
                    if name == 'DisplayConnections':
                        EasyshellLib.getElement('UserTitles').Click()
                        time.sleep(1)
                        if not EasyshellLib.getElement('UserConnection').IsOffScreen:
                            self.Logfile("[PASS]: {} is shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is not shown".format(name))
                            self.capture(profile, "[Fail]: {} is not shown".format(name))
                    if name == 'DisplayStoreFront':
                        EasyshellLib.getElement('UserTitles').Click()
                        time.sleep(1)
                        if not EasyshellLib.getElement('UserStoreFront').IsOffScreen:
                            self.Logfile("[PASS]: {} is shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is not shown".format(name))
                            self.capture(profile, "[Fail]: {} is not shown".format(name))
                    if name == 'DisplayWebsites':
                        EasyshellLib.getElement('UserTitles').Click()
                        time.sleep(1)
                        if not EasyshellLib.getElement('UserWebsites').IsOffScreen:
                            self.Logfile("[PASS]: {} is shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is not shown".format(name))
                            self.capture(profile, "[Fail]: {} is not shown".format(name))
                    # ////////item for Web browser ///////////////////////////////
                    # ---------No navigation buttons---
                    # ---------No
                    if name == 'DisplayAddress':
                        EasyshellLib.getElement('UserBrowser').Click()
                        time.sleep(1)
                        if not EasyshellLib.getElement('AddressBar').IsOffScreen:
                            self.Logfile("[PASS]: {} is shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is not shown".format(name))
                            self.capture(profile, "[Fail]: {} is not shown".format(name))
                    if name == 'DisplayHome':
                        EasyshellLib.getElement('UserBrowser').Click()
                        time.sleep(1)
                        if not EasyshellLib.getElement('WebHome').IsOffScreen:
                            self.Logfile("[PASS]: {} is shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is not shown".format(name))
                            self.capture(profile, "[Fail]: {} is not shown".format(name))
                    # ------------------------Item for Admin power----------------------------------------
                    if name == 'AllowLock':
                        EasyshellLib.getElement('UserPower').Click()
                        time.sleep(1)
                        if not EasyshellLib.getElement('Lock').IsOffScreen:
                            self.Logfile("[PASS]: {} is shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is not shown".format(name))
                            self.capture(profile, "[Fail]: {} is not shown".format(name))
                    if name == 'AllowLogoff':
                        EasyshellLib.getElement('UserPower').Click()
                        time.sleep(1)
                        if not EasyshellLib.getElement('Logoff').IsOffScreen:
                            self.Logfile("[PASS]: {} is shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is not shown".format(name))
                            self.capture(profile, "[Fail]: {} is not shown".format(name))
                    if name == 'AllowRestart':
                        EasyshellLib.getElement('UserPower').Click()
                        time.sleep(1)
                        if not EasyshellLib.getElement('Restart').IsOffScreen:
                            self.Logfile("[PASS]: {} is shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is not shown".format(name))
                            self.capture(profile, "[Fail]: {} is not shown".format(name))
                    if name == 'AllowShutDown':
                        EasyshellLib.getElement('UserPower').Click()
                        time.sleep(1)
                        if not EasyshellLib.getElement('Shutdown').IsOffScreen:
                            self.Logfile("[PASS]: {} is shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is not shown".format(name))
                            self.capture(profile, "[Fail]: {} is not shown".format(name))
                    # ----------------- Virtual keyboard --------
                    # -------------- No legacy touch keyboard---------
                    if name == 'DisplayVKeyboard':
                        if not EasyshellLib.getElement('UserKeyBoard').IsOffScreen:
                            self.Logfile("[PASS]: {} is shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is not shown".format(name))
                            self.capture(profile, "[Fail]: {} is not shown".format(name))
                    # ---Label that display mac/time/version... at the bottom of UI -----
                    if name == 'DisplayTime':
                        if not EasyshellLib.getElement('Time').IsOffScreen:
                            self.Logfile("[PASS]: {} is shown".format(name))
                            real_time = EasyshellLib.CommonUtils.getLocalTime('%H:%M')
                            show_time = EasyshellLib.getElement('Time').Name
                            if real_time.split(':')[0] in show_time or str(
                                    int(real_time.split(':')[0]) + 12) in show_time:
                                self.Logfile("-->[PASS]: {} real time format".format(name))
                            else:
                                self.Logfile("-->[Fail]: {} real time format".format(name))
                                self.capture(profile, "-->[Fail]: {} real time format".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is not shown".format(name))
                            self.capture(profile, "[Fail]: {} is not shown".format(name))
                    if name == 'DisplayIP':
                        if not EasyshellLib.getElement('IPAddr').IsOffScreen:
                            self.Logfile("[PASS]: {} is shown".format(name))
                            print(EasyshellLib.CommonUtils.getNetInfo(), '--------net info')
                            real_ip = EasyshellLib.CommonUtils.getNetInfo()['IP']
                            show_ip = EasyshellLib.getElement('IPAddr').Name
                            if real_ip == show_ip:
                                self.Logfile("-->[PASS]: {} real IP".format(name))
                            else:
                                self.Logfile("-->[Fail]: {} real IP".format(name))
                                self.capture(profile, "-->[Fail]: {} real IP".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is not shown".format(name))
                            self.capture(profile, "[Fail]: {} is not shown".format(name))
                    if name == 'DisplayMAC':
                        if not EasyshellLib.getElement('MACAddr').IsOffScreen:
                            self.Logfile("[PASS]: {} is shown".format(name))
                            real_mac = EasyshellLib.CommonUtils.getNetInfo()['MAC']
                            show_mac = EasyshellLib.getElement('MACAddr').Name
                            if real_mac == show_mac:
                                self.Logfile("-->[PASS]: {} real mac".format(name))
                            else:
                                self.Logfile("-->[Fail]: {} real mac".format(name))
                                self.capture(profile, "-->[Fail]: {} real mac".format(name))

                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is not shown".format(name))
                            self.capture(profile, "[Fail]: {} is not shown".format(name))
                    # ------------No Enable Network status notification -----------
                    # ------------No Hide HP Easy Shell during Session ------------
                    # -----------No Custom Background -----------------------------
                elif status == "OFF":
                    # ///////判断主按钮是否关闭，如果OFF,子按钮不再检查
                    if not swTitles:
                        if name in ['DisplayApp', 'DisplayConnections', 'DisplayStoreFront', 'DisplayWebsites']:
                            continue
                    if not swBrowser:
                        if name in ['DisplayAddress', 'DisplayNavigation', 'DisplayHome']:
                            continue
                    if not swPower:
                        if name in ['AllowLock', 'AllowLogoff', 'AllowRestart', 'AllowShutDown']:
                            continue
                    # ///////////////////////////////////
                    if name == 'DisplayTitle':
                        if EasyshellLib.getElement('UserTitles').IsOffScreen:
                            swTitles = False
                            self.Logfile("[PASS]: {} is not shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is shown".format(name))
                            self.capture(profile, "[Fail]: {} is shown".format(name))
                    if name == 'DisplayBrowser':
                        if EasyshellLib.getElement('UserBrowser').IsOffScreen:
                            swBrowser = False
                            self.Logfile("[PASS]: {} is not shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is shown".format(name))
                            self.capture(profile, "[Fail]: {} is shown".format(name))
                    if name == 'DisplayAdmin':
                        if EasyshellLib.getElement('UserAdmin').IsOffScreen:
                            self.Logfile("[PASS]: {} is not shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is shown".format(name))
                            self.capture(profile, "[Fail]: {} is shown".format(name))
                    if name == 'DisplayPower':
                        if EasyshellLib.getElement('UserPower').IsOffScreen:
                            swPower = False
                            self.Logfile("[PASS]: {} is not shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is shown".format(name))
                            self.capture(profile, "[Fail]: {} is shown".format(name))
                    # ////////// item for Titles //////////////////////////////////////
                    if name == 'DisplayApp':
                        EasyshellLib.getElement('UserTitles').Click()
                        time.sleep(1)
                        if EasyshellLib.getElement('UserApp').IsOffScreen:
                            self.Logfile("[PASS]: {} is not shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is shown".format(name))
                            self.capture(profile, "[Fail]: {} is shown".format(name))
                    if name == 'DisplayConnections':
                        EasyshellLib.getElement('UserTitles').Click()
                        time.sleep(1)
                        if EasyshellLib.getElement('UserConnection').IsOffScreen:
                            self.Logfile("[PASS]: {} is not shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is shown".format(name))
                            self.capture(profile, "[Fail]: {} is shown".format(name))
                    if name == 'DisplayStoreFront':
                        EasyshellLib.getElement('UserTitles').Click()
                        time.sleep(1)
                        if EasyshellLib.getElement('UserStoreFront').IsOffScreen:
                            self.Logfile("[PASS]: {} is not shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is shown".format(name))
                            self.capture(profile, "[Fail]: {} is shown".format(name))
                    if name == 'DisplayWebsites':
                        EasyshellLib.getElement('UserTitles').Click()
                        time.sleep(1)
                        if EasyshellLib.getElement('UserBrowser').IsOffScreen:
                            self.Logfile("[PASS]: {} is not shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is shown".format(name))
                            self.capture(profile, "[Fail]: {} is shown".format(name))
                    # ////////item for Web browser ///////////////////////////////
                    if name == 'DisplayAddress':
                        EasyshellLib.getElement('UserBrowser').Click()
                        time.sleep(1)
                        if EasyshellLib.getElement('AddressBar').IsOffScreen:
                            self.Logfile("[PASS]: {} is not shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is shown".format(name))
                            self.capture(profile, "[Fail]: {} is shown".format(name))
                    if name == 'DisplayHome':
                        EasyshellLib.getElement('UserBrowser').Click()
                        time.sleep(1)
                        if EasyshellLib.getElement('WebHome').IsOffScreen:
                            self.Logfile("[PASS]: {} is not shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is shown".format(name))
                            self.capture(profile, "[Fail]: {} is shown".format(name))
                    # ------------------------Item for Admin power----------------------------------------
                    if name == 'AllowLock':
                        EasyshellLib.getElement('UserPower').Click()
                        time.sleep(1)
                        if EasyshellLib.getElement('Lock').IsOffScreen:
                            self.Logfile("[PASS]: {} is not shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is shown".format(name))
                            self.capture(profile, "[Fail]: {} is shown".format(name))
                    if name == 'AllowLogoff':
                        EasyshellLib.getElement('UserPower').Click()
                        time.sleep(1)
                        if EasyshellLib.getElement('Logoff').IsOffScreen:
                            self.Logfile("[PASS]: {} is not shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is shown".format(name))
                            self.capture(profile, "[Fail]: {} is shown".format(name))
                    if name == 'AllowRestart':
                        EasyshellLib.getElement('UserPower').Click()
                        time.sleep(1)
                        if EasyshellLib.getElement('Restart').IsOffScreen:
                            self.Logfile("[PASS]: {} is not shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is shown".format(name))
                            self.capture(profile, "[Fail]: {} is shown".format(name))
                    if name == 'AllowShutDown':
                        EasyshellLib.getElement('UserPower').Click()
                        time.sleep(1)
                        if EasyshellLib.getElement('Shutdown').IsOffScreen:
                            self.Logfile("[PASS]: {} is not shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is shown".format(name))
                            self.capture(profile, "[Fail]: {} is shown".format(name))
                    # ----------------- Virtual keyboard --------
                    if name == 'DisplayVKeyboard':
                        if EasyshellLib.getElement('UserKeyBoard').IsOffScreen:
                            self.Logfile("[PASS]: {} is not shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is shown".format(name))
                            self.capture(profile, "[Fail]: {} is shown".format(name))
                    # ---Label that display mac/time/version... at the bottom of UI -----
                    if name == 'DisplayTime':
                        MainWnd = EasyshellLib.getElement("MAIN_WINDOW")
                        if EasyshellLib.getElement('Time', parent=MainWnd).IsOffScreen:
                            self.Logfile("[PASS]: {} is not shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is shown".format(name))
                            self.capture(profile, "[Fail]: {} is shown".format(name))
                    if name == 'DisplayIP':
                        if EasyshellLib.getElement('IPAddr').IsOffScreen:
                            self.Logfile("[PASS]: {} is not shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is shown".format(name))
                            self.capture(profile, "[Fail]: {} is shown".format(name))
                    if name == 'DisplayMAC':
                        if EasyshellLib.getElement('MACAddr').IsOffScreen:
                            self.Logfile("[PASS]: {} is not shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is shown".format(name))
                            self.capture(profile, "[Fail]: {} is shown".format(name))
                    # ------------No Enable Network status notification -----------
                    # ------------No Hide HP Easy Shell during Session ------------
            except:
                flag = False
                self.Logfile("[Exception]: {}\n{}".format(name, traceback.format_exc()))
                self.capture(profile, "[Exception]: {}\n{}".format(name, traceback.format_exc()))
        return flag


class UserSettings(EasyShellTest):
    def __init__(self):
        EasyShellTest.__init__(self)
        self.section_name = 'userSettings'

    def edit(self, profile):
        """
        :param profile: one of test parameters' combination
        """
        flag = True  # record the function's status
        test = self.sections[self.section_name][profile]
        self.launch()
        self.Logfile('---------Begin To Edit User settings----------')
        EasyshellLib.getElement('Settings').Click()
        EasyshellLib.getElement('KioskMode').Enable()
        for item in test:
            name = item.split(":")[0].strip()  # setting name
            status = item.split(":")[1].strip()  # setting status on/off
            if status.upper() == 'ON':
                try:
                    EasyshellLib.getElement(name).Enable()
                except:
                    flag = False
                    self.Logfile('[Fail]Button {} Enable\n{}'.format(name, traceback.format_exc()))
                    self.capture(profile, '[Fail]Button {} Enable\n{}'.format(name, traceback.format_exc()))
            elif status.upper() == 'OFF':
                try:
                    EasyshellLib.getElement(name).Disable()
                except:
                    flag = False
                    self.Logfile('[Fail]Button {} Disable\n{}'.format(name, traceback.format_exc()))
                    self.capture(profile, '[Fail]Button {} Disable\n{}'.format(name, traceback.format_exc()))
            else:
                flag = False
                self.Logfile('[Fail]Button {} Status in test data is not Correct!'.format(name))
                self.capture(profile, '[Fail]Button {} Status in test data is not Correct!'.format(name))
        EasyshellLib.getElement('APPLY').Click()
        EasyshellLib.getElement('Exit').Click()
        self.Logfile('[PASS] Modify user settings')
        return flag

    def check(self, profile):
        flag = True
        # ---Below test use yml, list type can be test in order --------
        test = self.sections[self.section_name][profile]
        # 以下为部分设置的总开关sw = switch
        self.Logfile('---------Begin To Check User settings----------')
        if EasyshellLib.getElement('MAIN_WINDOW').Exists():
            EasyshellLib.getElement('MAIN_WINDOW').SetFocus()

        if test[0].split(":")[1].strip() == "OFF":
            if EasyshellLib.getElement('UserSettings').IsOffScreen:
                self.Logfile('[PASS]:Allow User settings is OFF, user setting is not shown')
                return True
            else:
                self.Logfile('[FAIL]:Allow User settings is OFF, But user setting is shown')
                self.capture(profile, "[Fail]: Allow User settings is OFF, But user setting is shown")
                return False
        EasyshellLib.getElement('UserSettings').Click()
        for item in test:
            """
            1. 没有测试总开关OFF,子开关ON时的状态
            2. Wifi 只测试链接窗口是否弹出，没有测试wifi连接功能
            """
            name = item.split(":")[0].strip()
            status = item.split(":")[1].strip()
            if status == "ON":
                # ------------User Settings -----------------------------------
                if name == 'AllowMouse':
                    if not EasyshellLib.getElement('SysMouseIcon').IsOffScreen:
                        EasyshellLib.getElement('SysMouseIcon').Click()
                        time.sleep(2)
                        if EasyshellLib.getElement('MousedSetting').Exists():
                            EasyshellLib.getElement('MousedSetting').Close()
                            self.Logfile("[PASS]: {} Can be launch".format(name))
                        else:
                            flag = False
                            self.capture(profile,
                                         "[Fail]: {} can not be launch, MousedSetting dialog not exist".format(name))
                            self.Logfile("[Fail]: {} can not be launch, MousedSetting dialog not exist".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                        self.capture(profile, "[Fail]: {} is not shown".format(name))
                if name == 'AllowKeyboard':
                    if not EasyshellLib.getElement('SysKeyboardIcon').IsOffScreen:
                        EasyshellLib.getElement('SysKeyboardIcon').Click()
                        if EasyshellLib.getElement('KeyboardSetting').Exists(3, 0):
                            EasyshellLib.getElement('KeyboardSetting').Close()
                            self.Logfile("[PASS]: {} is shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} can not be launch".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                        self.capture(profile, "[Fail]: {} is not shown".format(name))
                if name == 'AllowDisplay':
                    logon_user = platform.version()
                    if logon_user.split(".")[0] == "10":
                        if not EasyshellLib.getElement('SysDisplayIcon').IsOffScreen:
                            EasyshellLib.getElement('SysDisplayIcon').Click()
                            if EasyshellLib.getElement('DisplaySetting').Exists(3, 0):
                                EasyshellLib.getElement('DisplaySetting').Close()
                                self.Logfile("[PASS]: {} is shown".format(name))
                            else:
                                flag = False
                                self.Logfile("[Fail]: {} is not shown".format(name))
                                self.capture(profile, "[Fail]: {} is not shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is not shown".format(name))
                            self.capture(profile, "[Fail]: {} is not shown".format(name))
                    else:
                        if not EasyshellLib.getElement('SysDisplayIcon').IsOffScreen:
                            EasyshellLib.getElement('SysDisplayIcon').Click()
                            if EasyshellLib.getElement('DisplaySetting_7').Exists(3, 0):
                                EasyshellLib.getElement('DisplaySetting_7').Close()
                                self.Logfile("[PASS]: {} is shown".format(name))
                            else:
                                flag = False
                                self.Logfile("[Fail]: {} is not shown".format(name))
                                self.capture(profile, "[Fail]: {} is not shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is not shown".format(name))
                            self.capture(profile, "[Fail]: {} is not shown".format(name))
                if name == 'AllowSound':
                    if not EasyshellLib.getElement('SysSoundIcon').IsOffScreen:
                        EasyshellLib.getElement('SysSoundIcon').Click()
                        if EasyshellLib.getElement('SoundSetting').Exists(3, 0):
                            EasyshellLib.getElement('SoundSetting').Close()
                            self.Logfile("[PASS]: {} is shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is not shown".format(name))
                            self.capture(profile, "[Fail]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                        self.capture(profile, "[Fail]: {} is not shown".format(name))
                if name == 'AllowRegion':
                    logon_user = platform.version()
                    if logon_user.split(".")[0] == "10":
                        # _os = "WES10"
                        if not EasyshellLib.getElement('SysRegionIcon').IsOffScreen:
                            EasyshellLib.getElement('SysRegionIcon').Click()
                            if EasyshellLib.getElement('RegionSetting').Exists(3, 0):
                                EasyshellLib.getElement('RegionSetting').Close()
                                self.Logfile("[PASS]: {} is shown".format(name))
                            else:
                                flag = False
                                self.Logfile("[Fail]: {} is not shown".format(name))
                                self.capture(profile, "[Fail]: {} is not shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is not shown".format(name))
                            self.capture(profile, "[Fail]: {} is not shown".format(name))
                    else:
                        # _os = "WES7"
                        if not EasyshellLib.getElement('SysRegionIcon').IsOffScreen:
                            EasyshellLib.getElement('SysRegionIcon').Click()
                            if EasyshellLib.getElement('RegionSetting_7').Exists(3, 0):
                                EasyshellLib.getElement('RegionSetting_7').Close()
                                self.Logfile("[PASS]: {} is shown".format(name))
                            else:
                                flag = False
                                self.Logfile("[Fail]: {} is not shown".format(name))
                                self.capture(profile, "[Fail]: {} is not shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is not shown".format(name))
                            self.capture(profile, "[Fail]: {} is not shown".format(name))
                if name == 'AllowNetworkConn':
                    if not EasyshellLib.getElement('SysNetworkConnIcon').IsOffScreen:
                        EasyshellLib.getElement('SysNetworkConnIcon').Click()
                        if EasyshellLib.getElement('NetworkSetting').Exists(3, 0):
                            EasyshellLib.getElement('NetworkSetting').Close()
                            self.Logfile("[PASS]: {} is shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is not shown".format(name))
                            self.capture(profile, "[Fail]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                        self.capture(profile, "[Fail]: {} is not shown".format(name))
                if name == 'AllowDateTime':
                    if not EasyshellLib.getElement('SysDateTimeIcon').IsOffScreen:
                        EasyshellLib.getElement('SysDateTimeIcon').Click()
                        if EasyshellLib.getElement('DateTimeSetting').Exists(3, 0):
                            EasyshellLib.getElement('DateTimeSetting').Close()
                            self.Logfile("[PASS]: {} is shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is not shown".format(name))
                            self.capture(profile, "[Fail]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                        self.capture(profile, "[Fail]: {} is not shown".format(name))
                if name == 'AllowEasyAccess':
                    if not EasyshellLib.getElement('SysEaseAccessCenterIcon').IsOffScreen:
                        EasyshellLib.getElement('SysEaseAccessCenterIcon').Click()
                        if EasyshellLib.getElement('EaseAccessSetting').Exists(3, 0):
                            EasyshellLib.getElement('EaseAccessSetting').Close()
                            self.Logfile("[PASS]: {} is shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is not shown".format(name))
                            self.capture(profile, "[Fail]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                        self.capture(profile, "[Fail]: {} is not shown".format(name))
                if name == 'AllowIEProperty':
                    if not EasyshellLib.getElement('SysIEIcon').IsOffScreen:
                        EasyshellLib.getElement('SysIEIcon').Click()
                        if EasyshellLib.getElement('InternetSetting').Exists(3, 0):
                            EasyshellLib.getElement('InternetSetting').Close()
                            self.Logfile("[PASS]: {} is shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is not shown".format(name))
                            self.capture(profile, "[Fail]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                        self.capture(profile, "[Fail]: {} is not shown".format(name))
                if name == 'AllowWifiConfig':
                    if not EasyshellLib.getElement('SysWirelessIcon').IsOffScreen:
                        EasyshellLib.getElement('SysWirelessIcon').Click()
                        if EasyshellLib.getElement('WirelessSetting').Exists(3, 0):
                            EasyshellLib.getElement('WirelessSetting').Close()
                            self.Logfile("[PASS]: {} is shown".format(name))
                        else:
                            flag = False
                            self.Logfile("[Fail]: {} is not shown".format(name))
                            self.capture(profile, "[Fail]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                        self.capture(profile, "[Fail]: {} is not shown".format(name))
            elif status == "OFF":
                if name == 'AllowMouse':
                    if not EasyshellLib.getElement('SysMouseIcon').Exists(0, 0):
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                        self.capture(profile, "[Fail]: {} is shown".format(name))
                if name == 'AllowKeyboard':
                    if not EasyshellLib.getElement('SysKeyboardIcon').Exists(0, 0):
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                        self.capture(profile, "[Fail]: {} is shown".format(name))
                if name == 'AllowDisplay':
                    if not EasyshellLib.getElement('SysDisplayIcon').Exists(0, 0):
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                        self.capture(profile, "[Fail]: {} is shown".format(name))
                if name == 'AllowSound':
                    if not EasyshellLib.getElement('SysSoundIcon').Exists(0, 0):
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                        self.capture(profile, "[Fail]: {} is shown".format(name))
                if name == 'AllowRegion':
                    if not EasyshellLib.getElement('SysRegionIcon').Exists(0, 0):
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                        self.capture(profile, "[Fail]: {} is shown".format(name))
                if name == 'AllowNetworkConn':
                    if not EasyshellLib.getElement('SysNetworkConnIcon').Exists(0, 0):
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                        self.capture(profile, "[Fail]: {} is shown".format(name))
                if name == 'AllowDateTime':
                    if not EasyshellLib.getElement('SysDateTimeIcon').Exists(0, 0):
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                        self.capture(profile, "[Fail]: {} is shown".format(name))
                if name == 'AllowEasyAccess':
                    if not EasyshellLib.getElement('SysEaseAccessCenterIcon').Exists(0, 0):
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                        self.capture(profile, "[Fail]: {} is shown".format(name))
                if name == 'AllowIEProperty':
                    if not EasyshellLib.getElement('SysIEIcon').Exists(0, 0):
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                        self.capture(profile, "[Fail]: {} is shown".format(name))
                if name == 'AllowWifiConfig':
                    if not EasyshellLib.getElement('SysWirelessIcon').Exists(0, 0):
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                        self.capture(profile, "[Fail]: {} is shown".format(name))
        return flag


class Background(EasyShellTest):
    def __init__(self):
        EasyShellTest.__init__(self)

    # -------------------------------back ground ---------------------------------------
    def set_background(self, bg='Custom'):
        # bg support:
        # {custom | default | Blue | Green | Red | Purple | light | dark}
        # please strictly use above value, including upper and lower letters
        self.Logfile('---------Begin To Set Background ---------------')
        UserSettings().modify('test1')
        self.launch()
        try:
            print(bg)
            if bg == 'Custom':
                file = 'customBG'
                EasyshellLib.getElement('AllowUserSetting').Enable()
                EasyshellLib.getElement('EnableCustom').Enable()
                EasyshellLib.getElement('BGFileLocationButton').GetInvokePattern().Invoke()
                EasyshellLib.getElement('BGFileURLEdit').SetValue(os.path.join(self.misc, "%s.jpg" % file))
                EasyshellLib.getElement('BGFileOpen').Click()
                EasyshellLib.getElement('APPLY').Click()
                self.Logfile('[PASS]: Set {} Background file\n'.format(bg))
                return True
            else:
                userSettings = self.sections['userSettings']['test1']
                for item in userSettings:
                    EasyshellLib.getElement(item.split(':')[0]).Enable()
                # EasyshellLib.getElement('AllowUserSetting').Enable()

                # for t in EasyshellLib.UserSettings_Dict.values():
                #     t.Enable()
                EasyshellLib.getElement('EnableCustom').Disable()
                EasyshellLib.getElement('SelectTheme').GetInvokePattern().Invoke()
                # -------------------select listitem by match the name----------------
                bgcomb = EasyshellLib.getElement('BGTheme')
                bgcomb.Click()
                time.sleep(3)
                txt = CommonLib.TextControl(RegexName='.*%s.*' % bg)
                txt.Click()
                EasyshellLib.getElement('OK').GetInvokePattern().Invoke()
                EasyshellLib.getElement('APPLY').Click()
                self.Logfile('[PASS]: Set {} Background file\n'.format(bg))
                return True
        except:
            self.Logfile('[FAIL]: Set {} Background file\n{}\n'.format(bg, traceback.format_exc()))
            self.capture("background", '[FAIL]: Set {} Background file\n{}\n'.format(bg, traceback.format_exc()))
            return False

    def check_background(self, bg='custom'):
        self.Logfile('---------Begin To Check Background ---------------')
        EasyshellLib.getElement('UserSettings').Click()
        CommonLib.PaneControl(AutomationId='mainFrame').CaptureToImage('%s.jpg' % bg)
        rgb1 = EasyshellLib.CommonUtils.getPicRGB('%s.jpg' % bg)
        rgb2 = EasyshellLib.CommonUtils.getPicRGB(os.path.join(self.misc, '%s.jpg' % bg))
        compare = EasyshellLib.CommonUtils.compareByRGB(rgb1, rgb2)
        print(compare)
        os.remove('%s.jpg' % bg)
        if compare > 0.9:
            print('pass')
            self.Logfile('[PASS]: check {} Background file\n'.format(bg))
            return True
        else:
            print('Fail')
            self.Logfile('[FAIL]: Check {} Background file\n'.format(bg))
            self.capture("background", '[FAIL]: Check {} Background file\n'.format(bg))
            return False


class Shell_Application(EasyShellTest):
    def __init__(self):
        EasyShellTest.__init__(self)
        self.section_name = 'createApp'

    # --------------- Applications Creation --------------------
    @pysnooper.snoop(EasyShellTest().debug)
    def check(self, profile):
        self.Logfile('-------------Begin to Check Application --------------')
        flag = True
        test = self.sections[self.section_name][profile]
        Name = test["Name"]
        Launchdelay = test['Launchdelay']
        Autolaunch = test['Autolaunch']
        Persistent = test['Persistent']
        Maximized = test['Maximized']
        AdminOnly = test['Adminonly']
        HideMissApp = test['Hidemissapp']
        WindowName = test['WindowName']
        try:
            if EasyshellLib.getElement('MAIN_WINDOW').Exists():
                EasyshellLib.getElement('MAIN_WINDOW').SetFocus()
            EasyshellLib.getElement('UserTitles').Click()
            if AdminOnly:
                if self.utils(profile, 'shown'):
                    flag = False
                    self.Logfile("[Fail]:App {} AdminOnly: {}, {} is shown".format(Name, AdminOnly, profile))
                    self.capture(profile, "[Fail]:App {} AdminOnly: {}, {} is shown".format(Name, AdminOnly, profile))
                    return flag
                else:
                    flag = True
                    self.Logfile("[PASS]:App {} AdminOnly:{}".format(Name, AdminOnly))
                    return flag
            else:
                if not self.utils(profile, 'shown'):
                    flag = False
                    self.Logfile("[Fail]:App {} AdminOnly:{}, {} is not shown".format(Name, AdminOnly, profile))
                    self.capture(profile,
                                 "[Fail]:App {} AdminOnly:{}, {} is not shown".format(Name, AdminOnly, profile))
                    return flag
            if HideMissApp:
                if self.utils(profile, 'shown'):
                    flag = False
                    self.Logfile("[Failed]:App {} Hide Missing App, {} is shown".format(Name, profile))
                    self.capture(profile, "[Failed]:App {} Hide Missing App, {} is shown".format(Name, profile))
                    return flag
                else:
                    self.Logfile("[PASS]:App {} Hide Missing App".format(Name))
                    return flag
            if Autolaunch == 0:
                self.utils(profile, "launch")
                if Launchdelay == 0:
                    with open('debug.txt', 'a') as f:
                        f.write('launch delay 0:{}\n'.format(Launchdelay))
                    for t in range(5):
                        if CommonLib.WindowControl(RegexName=WindowName).Exists(0, 0):
                            break
                        else:
                            continue
                    if not CommonLib.WindowControl(RegexName=WindowName).Exists(0, 0):
                        flag = False
                        self.Logfile("[Failed]:App {} Manual Launch, Window {} not shown".format(Name, WindowName))
                        self.capture(profile,
                                     "[Failed]:App {} Manual Launch, Window {} not shown".format(Name, WindowName))
                        return flag
                else:
                    time.sleep(5)
                    if CommonLib.WindowControl(RegexName=WindowName).Exists(0, 0):
                        self.Logfile(
                            "[Failed]:APP {} Launch Delay, Expect App windows {} not exist".format(Name, WindowName))
                        self.capture(profile,
                                     "[Failed]:APP {} Launch Delay, Expect App {} windows not exist".format(Name,
                                                                                                            WindowName))
                        flag = False
                    else:
                        self.Logfile("[PASS]:APP {} Launch Delay".format(Name))
                    time.sleep(Launchdelay + 5)  # add 5s waitting app's launch time
            else:
                time.sleep(Launchdelay)
                if not CommonLib.WindowControl(RegexName=WindowName).Exists():
                    flag = False
                    self.Logfile("[Failed]:App {} AutoLaunch, {} not Exist".format(Name, WindowName))
                    self.capture(profile, "[Failed]:App {} AutoLaunch, {} not Exist".format(Name, WindowName))
                    return flag
                else:
                    self.Logfile("[PASS]:APP {} AutoLaunch".format(Name))
                    self.Logfile("[PASS]:APP {} Launch Delay".format(Name))
            if Maximized:
                if CommonLib.WindowControl(RegexName=WindowName).IsMaximize():
                    self.Logfile("[PASS]:App {} Maximized".format(Name))
                else:
                    self.Logfile("[Failed]:App {} Maximized".format(Name))
                    self.capture(profile, "[Failed]:App {} Maximized".format(Name))
                    flag = False
            if Persistent:
                if CommonLib.WindowControl(RegexName=WindowName).Exists():
                    CommonLib.WindowControl(RegexName=WindowName).Close()
                else:
                    flag = False
                    self.Logfile("[Failed]:App {} Persistent, app {} is not launched".format(Name, WindowName))
                    self.capture(profile, "[Failed]:App {} Persistent, app {} is not launched".format(Name, WindowName))
                    return flag
                time.sleep(3)
                if int(Launchdelay) == 0:
                    if not CommonLib.WindowControl(RegexName=WindowName).Exists(0, 0):
                        flag = False
                        self.Logfile("[Failed]:App {} Persistent".format(Name))
                        self.capture(profile, "[Failed]:App {} Persistent".format(Name))
                        return flag
                    else:
                        self.Logfile("[PASS]:App {} Persistent".format(Name))
                        # self.Logfile("[PASS]:App {} Not AutoDelay".format(Name))
                else:
                    time.sleep(5)
                    if CommonLib.WindowControl(RegexName=WindowName).Exists(0, 0):
                        flag = False
                        self.Logfile("[Failed]:APP {} AutoDelay".format(Name))
                        self.capture(profile, "[Failed]:APP {} AutoDelay".format(Name))
                    else:
                        time.sleep(int(Launchdelay))
                        if CommonLib.WindowControl(RegexName=WindowName).Exists(0, 0):
                            flag = True
                            self.Logfile("[PASS]:APP {} AutoDelay Persistent".format(Name))
                        else:
                            self.Logfile("[Failed]:APP {} AutoDelay persistent,expect shown".format(Name))
                            self.capture(profile, "[Failed]:APP {} AutoDelay persistent, expect shown".format(Name))
            else:
                if CommonLib.WindowControl(RegexName=WindowName).Exists():
                    CommonLib.WindowControl(RegexName=WindowName).Close()
                else:
                    flag = False
                    self.Logfile("[Failed]:App {} not Persistent, app {} is not launched".format(Name, WindowName))
                    self.capture(profile,
                                 "[Failed]:App {} not Persistent, app {} is not launched".format(Name, WindowName))
                    return flag
                time.sleep(5)
                if int(Launchdelay) == 0:
                    if CommonLib.WindowControl(RegexName=WindowName).Exists(0, 0):
                        flag = False
                        self.Logfile("[Failed]:App {} No Persistent".format(Name))
                        self.capture(profile, "[Failed]:App {} No Persistent".format(Name))
                        return flag
                    else:
                        self.Logfile("[PASS]:App {} No Persistent".format(Name))
                else:
                    time.sleep(Launchdelay)
                    if CommonLib.WindowControl(RegexName=WindowName).Exists(0, 0):
                        flag = False
                        self.Logfile("[Fail]:APP {} AutoDelay No Persistent, Expect No shown".format(Name))
                        self.capture(profile, "[Failed]:APP {} AutoDelay No persistent, expect shown".format(Name))
                    else:
                        self.Logfile("[PASS]:APP {} AutoDelay NO persistent".format(Name))
            return flag
        except:
            self.Logfile("[Failed]:App {} App Check\nErrors:\n{}\n".format(Name, traceback.format_exc()))
            self.capture(profile, "[Failed]:App {} App Check\nErrors:\n{}\n".format(Name, traceback.format_exc()))
            return False

    @pysnooper.snoop(EasyShellTest().debug)
    def create(self, profile):
        """
        profile is a test configuration name, it is a dict for app options, load from easyshell_testdata.yaml
        """
        test = self.sections['createApp'][profile]
        Name = test["Name"]
        Paths = test['Path']
        Path = ''
        for path in Paths:
            if os.path.exists(path):
                Path = path
                break
            else:
                Path = "ErrorPath"
        argument = test['Argument']
        launchdelay = test['Launchdelay']
        autolaunch = test['Autolaunch']
        persistent = test['Persistent']
        maximized = test['Maximized']
        adminonly = test['Adminonly']
        hidemissapp = test['Hidemissapp']
        try:
            self.launch()
            EasyshellLib.getElement('Settings').Click()
            EasyshellLib.getElement('KioskMode').Enable()
            EasyshellLib.getElement('DisplayTitle').Enable()
            EasyshellLib.getElement('DisplayApp').Enable()
            EasyshellLib.getElement('Applications').Click()
            if self.utils(profile, 'Exist'):
                self.utils(profile, 'Delete')
            EasyshellLib.getElement('ApplicationAdd').Click()
            EasyshellLib.CommonUtils.Wait(5)
            CommonLib.SendKeys(Name)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKeys(Path)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if argument == "None":
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKeys(argument)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if launchdelay == "None":
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_DELETE)
                CommonLib.SendKeys(str(launchdelay))
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if autolaunch:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if persistent == 0:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if maximized == 0:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if adminonly == 0:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if hidemissapp == 0:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
            EasyshellLib.getElement('APPLY').Click()
            self.Logfile("[PASS]:App {} Create".format(Name))
            return True
        except:
            self.Logfile("[FAIL]:App {} Create\nErrors:\n{}\n".format(Name, traceback.format_exc()))
            self.capture(profile, "[FAIL]:App {} Create\nErrors:\n{}\n".format(Name, traceback.format_exc()))
            return False

    @pysnooper.snoop(EasyShellTest().debug)
    def edit(self, newProfile, oldProfile):
        try:
            old = self.sections['createApp'][oldProfile]
            oldPaths = old['Path']
            oldPath = ''
            for path in oldPaths:
                if os.path.exists(path):
                    oldPath = path
                    break
            if oldPath == '':
                oldPath = oldPaths[0]
            oldAutolaunch = old['Autolaunch']
            oldPersistent = old['Persistent']
            oldMaximized = old['Maximized']
            oldAdminonly = old['Adminonly']
            oldHidemissapp = old['Hidemissapp']
            new = self.sections['createApp'][newProfile]
            newName = new["Name"]
            newPaths = new['Path']
            newPath = ''
            for path in newPaths:
                if os.path.exists(path):
                    newPath = path
                    break
            if newPath == '':
                newPath = newPaths[0]
            newArgument = new['Argument']
            newAutolaunch = new['Autolaunch']
            newPersistent = new['Persistent']
            newMaximized = new['Maximized']
            newAdminonly = new['Adminonly']
            newHidemissapp = new['Hidemissapp']
            newLaunchdelay = new['Launchdelay']
            self.Logfile('----------Begin to Edit Application with new profile {} ---------'.format(newProfile))
            self.launch()
            EasyshellLib.getElement('Applications').Click()
            # Modify setting//////////////////////////////////////
            if self.utils(newProfile, 'exist'):
                self.utils(newProfile, 'delete')
            self.utils(oldProfile, 'edit')
            time.sleep(3)
            ClearContent()
            CommonLib.SendKeys(newName)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            ClearContent(len(oldPath) + 5)
            CommonLib.SendKeys(newPath)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            ClearContent()
            if newArgument == 'None':
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKeys(newArgument)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            ClearContent()
            if newLaunchdelay == "None":
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_DELETE)
                CommonLib.SendKeys(str(newLaunchdelay))
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if oldAutolaunch == newAutolaunch:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if oldPersistent == newPersistent:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if oldMaximized == newMaximized:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if oldAdminonly == newAdminonly:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if oldHidemissapp == newHidemissapp:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
            CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
            EasyshellLib.getElement('APPLY').Click()
            self.Logfile("[PASS]:App {} Edit\n".format(newName))
            return True
        except:
            self.Logfile("[FAIL]:App {} Edit\nErrors:\n{}\n".format(newProfile, traceback.format_exc()))
            self.capture("EditProfile", "[FAIL]:App {} Edit\nErrors:\n{}\n".format(newProfile, traceback.format_exc()))
            return False


class Shell_Websites(EasyShellTest):
    def __init__(self):
        EasyShellTest.__init__(self)
        self.section_name = 'createWebsites'

    #   ---------------- Website Creation --------------------------
    def create(self, profile):
        try:
            test = self.sections[self.section_name][profile]
            Name = test["Name"]
            Address = test['Address']
            DefaultHome = test['DefaultHome']
            UseIE = test['UseIE']
            IEFullScreen = test['IEFullScreen']
            EmbedIE = test['EmbedIE']
            AllCloseEmbedIE = test['AllCloseEmbedIE']
            try:
                self.launch()
                EasyshellLib.getElement('Settings').Click()
                EasyshellLib.getElement('KioskMode').Enable()
                EasyshellLib.getElement('DisplayTitle').Enable()
                EasyshellLib.getElement('DisplayWebsites').Enable()
                EasyshellLib.getElement('DisplayBrowser').Enable()
                EasyshellLib.getElement('DisplayAddress').Enable()
                EasyshellLib.getElement('DisplayNavigation').Enable()
                EasyshellLib.getElement('DisplayHome').Enable()
                EasyshellLib.getElement('AllowUserSetting').Enable()
                EasyshellLib.getElement('AllowDisplay').Enable()
                EasyshellLib.getElement('WebSites').Click()
                if self.utils(profile, 'Exist'):
                    self.utils(profile, 'Delete')
                EasyshellLib.getElement('WebsiteAdd').Click()
                time.sleep(3)
                print(Name)
                CommonLib.SendKeys(Name)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
                CommonLib.SendKeys(Address)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
                if UseIE:
                    CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                    CommonLib.SendKey(CommonLib.Keys.VK_TAB)
                    if IEFullScreen:
                        CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                        CommonLib.SendKey(CommonLib.Keys.VK_TAB)
                        if EmbedIE:
                            CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
                            if AllCloseEmbedIE:
                                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                                CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
                                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                            else:
                                # not allcloseembedid |use Id | IE Fullscreen | EmbedIE
                                CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
                                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                        else:
                            # not embedie | use Ie | Full screen
                            CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
                            CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                    else:
                        # not IE fullscreen | use IE
                        CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
                        CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                else:
                    # Not use IE
                    CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
                    CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                if DefaultHome:
                    self.utils(profile, 'default')
                EasyshellLib.getElement('APPLY').Click()
                self.Logfile('[PASS] Create website {} test pass'.format(Name))
                return True
            except:
                self.Logfile("[FAIL]:Website {} Create\nErrors:\n{}\n".format(Name, traceback.format_exc()))
                self.capture('CreateWebsite',
                             "[FAIL]:Website {} Create\nErrors:\n{}\n".format(Name, traceback.format_exc()))
                return False
        except:
            self.Logfile("[FAIL]:Website {} Create\nErrors:\n{}\n".format(profile, traceback.format_exc()))
            self.capture('CreateWebsite',
                         "[FAIL]:Website {} Create\nErrors:\n{}\n".format(profile, traceback.format_exc()))
            return False

    # ---------------- Website Check --------------------------------------
    def check(self, profile):
        flag = True
        if EasyshellLib.getElement('MAIN_WINDOW').Exists():
            EasyshellLib.getElement('MAIN_WINDOW').SetFocus()
        test = self.sections[self.section_name][profile]
        DefaultHome = test['DefaultHome']
        UseIE = test['UseIE']
        IEFullScreen = test['IEFullScreen']
        EmbedIE = test['EmbedIE']
        AllCloseEmbedIE = test['AllCloseEmbedIE']
        EmbaedPaneName = test['EmbaedPaneName']
        if CommonLib.WindowControl(RegexName='.*- Internet Explorer').Exists(0, 0):
            CommonLib.WindowControl(RegexName='.*- Internet Explorer').Close()
        EasyshellLib.getElement('UserTitles').Click()
        if not self.utils(profile, 'launch'):
            self.Logfile('[Fail] check website {} error:Launch website fail'.format(profile))
            self.capture("CheckWeb", '[Fail] check website {} error:Launch website fail'.format(profile))
            return False
        time.sleep(5)
        if not UseIE:
            if CommonLib.PaneControl(RegexName=EmbaedPaneName).Exists(0, 0) and not EasyshellLib.getElement(
                    'AddressBar').IsOffScreen:
                if DefaultHome:
                    EasyshellLib.getElement('WebHome').Click()
                    time.sleep(5)
                    if CommonLib.PaneControl(RegexName=EmbaedPaneName).Exists(0, 0):
                        self.Logfile("[PASS]: Websites {} Check".format(profile))
                    else:
                        flag = False
                        self.Logfile("[FAIL]: Websites {} Check".format(profile))
                        self.capture("CheckWeb", "[FAIL]: Websites {} Check DefaultHOme".format(profile))
                else:
                    self.Logfile("[PASS]: Websites {} Check".format(profile))
            else:
                flag = False
                self.Logfile("[FAIL]: Websites {} Check".format(profile))
                self.capture("CheckWeb", "[FAIL]: Websites {} Check".format(profile))
        elif UseIE and not IEFullScreen:
            if CommonLib.WindowControl(RegexName=EmbaedPaneName).Exists(0, 0) and \
                    CommonLib.WindowControl(RegexName=EmbaedPaneName).PaneControl(AutomationId='41477').Exists():
                self.Logfile("[PASS]: Websites {} Check".format(profile))
            else:
                flag = False
                self.Logfile("[FAIL]: Websites {} Check".format(profile))
                self.capture("CheckWeb", "[FAIL]: Websites {} Check".format(profile))
        elif UseIE and IEFullScreen and not EmbedIE:
            if CommonLib.WindowControl(RegexName=EmbaedPaneName).Exists(0, 0) and \
                    not (
                            CommonLib.WindowControl(RegexName=EmbaedPaneName).PaneControl(AutomationId='41477').Exists(
                                0, 0)):
                self.Logfile("[PASS]: Websites {} Check".format(profile))
            else:
                flag = False
                self.Logfile("[FAIL]: Websites {} Check(110)".format(profile))
                self.capture("CheckWeb", "[FAIL]: Websites {} Check 110".format(profile))
        elif UseIE and IEFullScreen and EmbedIE and not AllCloseEmbedIE:
            if CommonLib.PaneControl(RegexName=EmbaedPaneName).Exists(0, 0) and \
                    EasyshellLib.getElement('AddressBar').IsOffScreen \
                    and EasyshellLib.getElement('WebIEClose').IsOffScreen:
                if DefaultHome:
                    EasyshellLib.getElement('WebHome').Click()
                    time.sleep(5)
                    if CommonLib.PaneControl(RegexName=EmbaedPaneName).Exists(0, 0):
                        self.Logfile("[PASS]: Websites {} Check".format(profile))
                    else:
                        flag = False
                        self.Logfile("[FAIL]: Websites {} Home Check(1110)".format(profile))
                        self.capture("CheckWeb", "[FAIL]: Websites {} Home Check(1110)".format(profile))
                else:
                    self.Logfile("[PASS]: Websites {} Check".format(profile))
            else:
                flag = False
                self.Logfile("[FAIL]: Websites {} Check(1110)".format(profile))
                self.capture("CheckWeb", "[FAIL]: Websites {} Check(1110)".format(profile))
        elif UseIE and IEFullScreen and EmbedIE and AllCloseEmbedIE:
            if CommonLib.PaneControl(RegexName=EmbaedPaneName).Exists(0, 0) and \
                    EasyshellLib.getElement('AddressBar').IsOffScreen \
                    and not EasyshellLib.getElement('WebIEClose').IsOffScreen:
                if DefaultHome:
                    EasyshellLib.getElement('WebHome').Click()
                    time.sleep(5)
                    if CommonLib.PaneControl(RegexName=EmbaedPaneName).Exists(0, 0):
                        self.Logfile("[PASS]: Websites {} Check".format(profile))
                    else:
                        flag = False
                        self.Logfile("[FAIL]: Websites {} Check(1111)".format(profile))
                        self.capture("CheckWeb", "[FAIL]: Websites {} Check(1111)".format(profile))
                else:
                    self.Logfile("[PASS]: Websites {} Check".format(profile))
            else:
                flag = False
                self.Logfile("[FAIL]: Websites {} Check(1111)".format(profile))
                self.capture("CheckWeb", "[FAIL]: Websites {} Check(1111)".format(profile))
        else:
            flag = False
            self.Logfile("[FAIL]: Websites {} Parameter Error!".format(profile))
            self.capture("CheckWeb", "[FAIL]: Websites {} Parameter Error!".format(profile))
        return flag

    def edit(self, newProfile, oldProfile):
        """
        Make sure Home Default is the same with OldProfile
        """
        if EasyshellLib.getElement('MAIN_WINDOW').Exists(0, 0):
            EasyshellLib.getElement('MAIN_WINDOW').Close()
        new = self.sections[self.section_name][newProfile]
        old = self.sections[self.section_name][oldProfile]
        newName = new["Name"]
        newAddress = new['Address']
        newUseIE = new['UseIE']
        oldUseIE = old['UseIE']
        newDefaultHome = new['DefaultHome']
        oldDefaultHome = old['DefaultHome']
        newIEFullScreen = new['IEFullScreen']
        oldIEFullScreen = old['IEFullScreen']
        newEmbedIE = new['EmbedIE']
        oldEmbedIE = old['EmbedIE']
        newAllCloseEmbedIE = new['AllCloseEmbedIE']
        oldAllCloseEmbedIE = old['AllCloseEmbedIE']
        try:
            self.Logfile('-----------Begin to Edit website -------------')
            self.launch()
            EasyshellLib.getElement('WebSites').Click()
            if self.utils(newProfile, 'exist'):
                self.utils(newProfile, 'delete')
            if not self.utils(oldProfile, 'exist'):
                self.Logfile('[Fail]: Edit website with new profile {}, old profile not exist'.format(newName))
                self.capture("EditWEB",
                             '[Fail]: Edit website with new profile {}, old profile not exist'.format(newName))
                return False
            self.utils(oldProfile, 'Edit')
            time.sleep(3)
            ClearContent()
            CommonLib.SendKeys(newName)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            ClearContent()
            CommonLib.SendKeys(newAddress)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if newUseIE != oldUseIE and newUseIE:
                if newUseIE:
                    CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                    CommonLib.SendKey(CommonLib.Keys.VK_TAB)
                else:
                    CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                    CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
                    CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                    if newDefaultHome != oldDefaultHome:
                        self.utils(newProfile, 'default')
                    EasyshellLib.getElement('APPLY').Click()
                    self.Logfile('[PASS] Edit website with new profile {}'.format(newName))
                    return True
            else:
                if newUseIE:
                    CommonLib.SendKey(CommonLib.Keys.VK_TAB)
                else:
                    CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
                    CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                    if newDefaultHome != oldDefaultHome:
                        self.utils(newProfile, 'default')
                    EasyshellLib.getElement('APPLY').Click()
                    self.Logfile('[PASS] Edit website with new profile {}'.format(newName))
                    return True
            if newIEFullScreen != oldIEFullScreen:
                if newIEFullScreen:
                    CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                    CommonLib.SendKey(CommonLib.Keys.VK_TAB)
                else:
                    CommonLib.SendKey(CommonLib.Keys.VK_TAB)
                    CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                    CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
                    CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                    if newDefaultHome != oldDefaultHome:
                        self.utils(newProfile, 'default')
                    EasyshellLib.getElement('APPLY').Click()
                    self.Logfile('[PASS] Edit website with new profile {}'.format(newName))
                    return True
            else:
                if newIEFullScreen:
                    CommonLib.SendKey(CommonLib.Keys.VK_TAB)
                else:
                    CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
                    CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                    if newDefaultHome != oldDefaultHome:
                        self.utils(newProfile, 'default')
                    EasyshellLib.getElement('APPLY').Click()
                    self.Logfile('[PASS] Edit website with new profile {}'.format(newName))
                    return True
            if newEmbedIE != oldEmbedIE:
                if newEmbedIE:
                    CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                    CommonLib.SendKey(CommonLib.Keys.VK_TAB)
                else:
                    CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                    CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
                    CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                    if newDefaultHome != oldDefaultHome:
                        self.utils(newProfile, 'default')
                    EasyshellLib.getElement('APPLY').Click()
                    self.Logfile('[PASS] Edit website with new profile {}'.format(newName))
                    return True
            else:
                if newEmbedIE:
                    CommonLib.SendKey(CommonLib.Keys.VK_TAB)
                else:
                    CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
                    CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                    if newDefaultHome != oldDefaultHome:
                        self.utils(newProfile, 'default')
                    EasyshellLib.getElement('APPLY').Click()
                    self.Logfile('[PASS] Edit website with new profile {}'.format(newName))
                    return True
            if newAllCloseEmbedIE != oldAllCloseEmbedIE:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                if newDefaultHome != oldDefaultHome:
                    self.utils(newProfile, 'default')
                EasyshellLib.getElement('APPLY').Click()
                self.Logfile('[PASS] Edit website with new profile {}'.format(newName))
                return True
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                if newDefaultHome != oldDefaultHome:
                    self.utils(newProfile, 'default')
                EasyshellLib.getElement('APPLY').Click()
                self.Logfile('[PASS] Edit website with new profile {}'.format(newName))
                return True
        except:
            self.Logfile(
                '[Fail]: Edit website with new profile {}, error: \n{}'.format(newName, traceback.format_exc()))
            self.capture("EditWEB",
                         '[Fail]: Edit website with new profile {}, error: \n{}'.format(newName,
                                                                                        traceback.format_exc()))
            return False


class Shell_StoreFront(EasyShellTest):
    def __init__(self):
        EasyShellTest.__init__(self)
        self.section_name = 'createStoreFront'

    # ----------------- StoreFront connection creation ----------------------------
    def create(self, profile):
        test = self.sections[self.section_name][profile]
        Name = test["Name"]
        URL = test['URL']
        SelectStore = test['SelectStore']
        StoreName = test['StoreName']
        Launchdelay = test['Launchdelay']
        LogonMethod = test['LogonMethod']
        Username = test['Username']
        Password = test['Password']
        Domain = test['Domain']
        HideDomain = test['HideDomain']
        CustomLogon = test['CustomLogon']
        Autolaunch = test['Autolaunch']
        ConnectionTimeout = test['ConnectionTimeout']
        DesktopToolbar = test['DesktopToolbar']
        try:
            self.launch()
            EasyshellLib.getElement('Settings').Click()
            EasyshellLib.getElement('KioskMode').Enable()
            EasyshellLib.getElement('DisplayTitle').Enable()
            EasyshellLib.getElement('DisplayStoreFront').Enable()
            EasyshellLib.getElement('StoreFront').Click()
            if self.utils(profile, 'Exist'):
                self.utils(profile, 'Delete')
            EasyshellLib.getElement('StoreFrontAdd').Click()
            time.sleep(5)
            CommonLib.SendKeys(Name)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKeys(URL)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
            CommonLib.SendKeys(LogonMethod)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if HideDomain == 'OFF':
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKeys(Username)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKeys(Password)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKeys(Domain)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if Autolaunch == 'OFF':
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if not SelectStore:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                time.sleep(1)
                CommonLib.SendKeys(StoreName)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                time.sleep(3)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=3)
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            ClearContent()
            CommonLib.SendKeys(str(Launchdelay))
            CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
            if CustomLogon == 'None':
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKeys(CustomLogon)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
            CommonLib.SendKey(CommonLib.Keys.VK_RIGHT)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=3)
            if DesktopToolbar == 'OFF':
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
            ClearContent()
            CommonLib.SendKeys(str(ConnectionTimeout))
            CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=4)
            CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
            EasyshellLib.getElement('APPLY').Click()
            self.Logfile('[PASS]: Storefront Connection {} Create'.format(Name))
            return True
        except Exception as e:
            self.Logfile("[FAIL]:Storfront {} Create\nErrors:\n{}\n".format(Name, e))
            return False

    def edit(self, newprofile, oldprofile):
        try:
            new = self.sections[self.section_name][newprofile]
            newName = new["Name"]
            newURL = new['URL']
            newSelectStore = new['SelectStore']
            newStoreName = new['StoreName']
            newLaunchdelay = new['Launchdelay']
            newLogonMethod = new['LogonMethod']
            newUsername = new['Username']
            newPassword = new['Password']
            newDomain = new['Domain']
            newHideDomain = new['HideDomain']
            newCustomLogon = new['CustomLogon']
            newAutolaunch = new['Autolaunch']
            newConnectionTimeout = new['ConnectionTimeout']
            newDesktopToolbar = new['DesktopToolbar']
            old = self.sections[self.section_name][oldprofile]
            oldName = old["Name"]
            oldURL = old['URL']
            oldSelectStore = old['SelectStore']
            oldStoreName = old['StoreName']
            oldLaunchdelay = old['Launchdelay']
            oldLogonMethod = old['LogonMethod']
            oldUsername = old['Username']
            oldPassword = old['Password']
            oldDomain = old['Domain']
            oldHideDomain = old['HideDomain']
            oldCustomLogon = old['CustomLogon']
            oldAutolaunch = old['Autolaunch']
            oldConnectionTimeout = old['ConnectionTimeout']
            oldDesktopToolbar = old['DesktopToolbar']
            self.launch()
            EasyshellLib.getElement('Settings').Click()
            EasyshellLib.getElement('KioskMode').Enable()
            EasyshellLib.getElement('DisplayTitle').Enable()
            EasyshellLib.getElement('DisplayStoreFront').Enable()
            EasyshellLib.getElement('StoreFront').Click()
            if self.utils(newprofile, 'Exist'):
                self.utils(newprofile, 'Delete')
            if self.utils(oldprofile, 'exist'):
                self.utils(oldprofile, 'edit')
            else:
                self.Logfile('[Fail]:Old StoreFont {} do not exist'.format(oldName))
                self.capture('EditStorefont', '[Fail]:Old StoreFont {} do not exist'.format(oldName))
                return False
            ClearContent(len(oldName) + 5)
            CommonLib.SendKeys(newName)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            ClearContent(len(oldURL) + 5)
            CommonLib.SendKeys(newURL)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
            CommonLib.SendKeys(newLogonMethod)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if newHideDomain == oldHideDomain:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            ClearContent(len(oldUsername) + 5)
            CommonLib.SendKeys(newUsername)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            ClearContent(len(oldPassword) + 5)
            CommonLib.SendKeys(newPassword)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            ClearContent(len(oldDomain) + 5)
            CommonLib.SendKeys(newDomain)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if newAutolaunch == oldAutolaunch:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
            ClearContent()
            CommonLib.SendKeys(str(newLaunchdelay))
            CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
            if newCustomLogon == oldCustomLogon:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                if newCustomLogon == 'None':
                    ClearContent(len(oldCustomLogon) + 5)
                    CommonLib.SendKey(CommonLib.Keys.VK_TAB)
                else:
                    CommonLib.SendKeys(newCustomLogon)
                    CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
            CommonLib.SendKey(CommonLib.Keys.VK_RIGHT)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=3)
            if newDesktopToolbar == oldDesktopToolbar:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
            ClearContent()
            CommonLib.SendKeys(str(newConnectionTimeout))
            CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=4)
            CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
            EasyshellLib.getElement('APPLY').Click()
            self.Logfile('[PASS]: View Connection {} Create'.format(newName))
            return True

        except:
            self.Logfile('[Failed]: RDP Connection {} Edit\n{}'.format(oldprofile, traceback.format_exc()))
            self.capture('EditRDP', '[Failed]: RDP Connection {} Edit\n{}'.format(oldprofile, traceback.format_exc()))
            return False

    def check(self, profile):
        if EasyshellLib.getElement('MAIN_WINDOW').Exists():
            EasyshellLib.getElement('MAIN_WINDOW').SetFocus()
        EasyshellLib.getElement('UserTitles').Click()
        if self.utils(profile, 'exist'):
            self.Logfile('[PASS]:Storefront connection {} Check Exist'.format(profile))
            return True
        else:
            self.Logfile('[FAIL]:Storefront connection {} Check Not Exist'.format(profile))
            self.capture('CheckRDP', '[FAIL]:Storefront connection {} Check Not Exist'.format(profile))
            return False


class Shell_View(EasyShellTest):
    def __init__(self):
        EasyShellTest.__init__(self)
        self.section_name = 'createView'

    # ----------------- VMWare view connection Creation ----------------------------
    @pysnooper.snoop(EasyShellTest().debug)
    def create(self, profile):
        test = self.sections['createView'][profile]
        Name = test["Name"]
        Hostname = test['Hostname']
        Launchdelay = test['Launchdelay']
        Argument = test['Argument']
        Autolaunch = test['Autolaunch']
        Persistent = test['Persistent']
        Layout = test['Layout']
        ConnUSBStartup = test['ConnUSBStartup']
        ConnUSBInsertion = test['ConnUSBInsertion']
        Username = test['Username']
        Password = test['Password']
        Domain = test['Domain']
        DesktopName = test['DesktopName']
        try:
            self.launch()
            EasyshellLib.getElement('KioskMode').Enable()
            EasyshellLib.getElement('DisplayTitle').Enable()
            EasyshellLib.getElement('DisplayConnections').Enable()
            EasyshellLib.getElement('Connections').Click()
            if self.utils(profile, 'exist', 'connection'):
                self.utils(profile, 'Delete', 'connection')
            EasyshellLib.getElement('VMwareAdd').Click()
            time.sleep(5)
            CommonLib.SendKeys(Name)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKeys(Hostname)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
            ClearContent()
            CommonLib.SendKeys(str(Launchdelay))
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if Argument == 'None':
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKeys(Argument)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if Autolaunch == 'OFF' or not Autolaunch:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if Persistent == 'OFF' or not Persistent:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
            CommonLib.SendKey(CommonLib.Keys.VK_RIGHT)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKeys(Layout)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if ConnUSBStartup == 'OFF':
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if ConnUSBInsertion == 'OFF':
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKeys(Username)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKeys(Password)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKeys(Domain)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKeys(DesktopName)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
            EasyshellLib.getElement('APPLY').Click()
            self.Logfile('[PASS]: View Connection {} Create'.format(Name))
            return True
        except:
            self.Logfile("[FAIL]: View Connection {} Create\nErrors:\n{}\n".format(Name, traceback.format_exc()))
            self.capture('CreateView',
                         "[FAIL]: View Connection {} Create\nErrors:\n{}\n".format(Name, traceback.format_exc()))
            return False

    @pysnooper.snoop(EasyShellTest().debug)
    def edit(self, newprofile, oldprofile):
        """
        copy from citrix, not modify
        """
        try:
            new = self.sections[self.section_name][newprofile]
            newname = new["Name"]
            newhostname = new['Hostname']
            newusername = new['Username']
            newpassword = new['Password']
            newdomain = new['Domain']
            newargument = new['Argument']
            newlayout = new['Layout']
            newconnUSBstartup = new['ConnUSBStartup']
            newconnUSBinsertion = new['ConnUSBInsertion']
            newdesktopname = new['DesktopName']
            newlaunchdelay = str(new['Launchdelay'])
            newautolaunch = new['Autolaunch']
            newpersistent = new['Persistent']
            old = self.sections[self.section_name][oldprofile]
            oldname = old["Name"]
            oldhostname = old['Hostname']
            oldusername = old['Username']
            olddomain = old['Domain']
            oldargument = old['Argument']
            oldconnUSBstartup = old['ConnUSBStartup']
            oldconnUSBinsertion = old['ConnUSBInsertion']
            olddesktopname = old['DesktopName']
            oldautolaunch = old['Autolaunch']
            oldpersistent = old['Persistent']
            self.launch()
            EasyshellLib.getElement('KioskMode').Enable()
            EasyshellLib.getElement('DisplayTitle').Enable()
            EasyshellLib.getElement('DisplayConnections').Enable()
            EasyshellLib.getElement('Connections').Click()
            if self.utils(newprofile, 'exist', 'connection'):
                self.utils(newprofile, 'delete', 'connection')
            if self.utils(oldprofile, 'exist', 'connection'):
                self.utils(oldprofile, 'Edit', 'connection')
            else:
                self.Logfile('[Fail]:Old View {} do not exist'.format(oldname))
                self.capture('EditView', '[Fail]:Old View {} do not exist'.format(oldname))
                return False
            time.sleep(5)
            ClearContent(len(oldname) + 5)
            CommonLib.SendKeys(newname)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            ClearContent(len(oldhostname) + 5)
            CommonLib.SendKeys(newhostname)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
            ClearContent(5)
            CommonLib.SendKeys(str(newlaunchdelay))
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            ClearContent(len(oldargument) + 5)
            CommonLib.SendKeys(newargument)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if newautolaunch == oldautolaunch:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if newpersistent == oldpersistent:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
            CommonLib.SendKey(CommonLib.Keys.VK_RIGHT)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKeys(newlayout)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if newconnUSBstartup == oldconnUSBstartup:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if newconnUSBinsertion == oldconnUSBinsertion:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            ClearContent(len(oldusername) + 5)
            CommonLib.SendKeys(newusername)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            ClearContent(30)
            CommonLib.SendKeys(newpassword)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            ClearContent(len(olddomain) + 5)
            CommonLib.SendKeys(newdomain)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            ClearContent(len(olddesktopname) + 5)
            CommonLib.SendKeys(newdesktopname)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
            EasyshellLib.getElement('APPLY').Click()
            self.Logfile('[PASS]: View Connection {} Edit'.format(oldname))
            return True
        except:
            self.Logfile('[Failed]: View Connection {} Edit\n{}'.format(oldprofile, traceback.format_exc()))
            self.capture("EditView", '[Failed]: View Connection {} Edit\n{}'.format(oldprofile, traceback.format_exc()))
            return False

    @pysnooper.snoop(EasyShellTest().debug)
    def check(self, profile):
        if EasyshellLib.getElement('MAIN_WINDOW').Exists():
            EasyshellLib.getElement('MAIN_WINDOW').SetFocus()
        EasyshellLib.getElement('UserTitles').Click()
        if self.utils(profile, 'exist', 'connection'):
            self.Logfile('[PASS]:View connection {} Check Exist'.format(profile))
            return True
        else:
            self.capture('CheckView', '[FAIL]:View connection {} Check Not Exist'.format(profile))
            self.Logfile('[FAIL]:View connection {} Check Not Exist'.format(profile))
            return False


class Shell_RDP(EasyShellTest):
    def __init__(self):
        EasyShellTest.__init__(self)
        self.section_name = 'createRDP'

    @pysnooper.snoop(EasyShellTest().debug)
    def create(self, profile):
        test = self.sections[self.section_name][profile]
        name = test["Name"]
        hostname = test['Hostname']
        username = test['Username']
        launchdelay = str(test['Launchdelay'])
        arguments = test['Arguments']
        autolaunch = test['Autolaunch']
        persistent = test['Persistent']
        customfile = test['Customfile']
        try:
            self.launch()
            EasyshellLib.getElement('Settings').Click()
            EasyshellLib.getElement('KioskMode').Enable()
            EasyshellLib.getElement('DisplayTitle').Enable()
            EasyshellLib.getElement('DisplayConnections').Enable()
            EasyshellLib.getElement('Connections').Click()
            if self.utils(profile, 'exist', 'connection'):
                self.utils(profile, 'Delete', 'connection')
            EasyshellLib.getElement('RDPAdd').Click()
            time.sleep(5)
            CommonLib.SendKeys(name)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKeys(hostname)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKeys(username)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            ClearContent(4)
            CommonLib.SendKeys(launchdelay)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if not arguments == 'None' or arguments is not None:
                CommonLib.SendKeys(arguments)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if autolaunch == 'OFF' or not autolaunch:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if persistent == 'OFF' or not persistent:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if customfile == 'OFF' or not customfile:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
                if os.path.exists(customfile):
                    CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                    time.sleep(1)
                    CommonLib.SendKeys(customfile)
                    CommonLib.SendKey(CommonLib.Keys.VK_ENTER)
                else:
                    self.Logfile('[Fail] Create RDP connection {}, customfile:{} not Exist'.format(name, customfile))
                    self.capture('CreateRDP',
                                 '[Fail] Create RDP connection {}, customfile:{} not Exist'.format(name, customfile))
                    CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
                    CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                    return False
                CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
            EasyshellLib.getElement('APPLY').Click()
            self.Logfile('[PASS]: Create RDP Connection {}'.format(name))
            return True
        except:
            self.Logfile('[Fail]: Create RDP Connection {}\n{}'.format(name, traceback.format_exc()))
            self.capture('CreateRDP', '[Fail]: Create RDP Connection {}\n{}'.format(name, traceback.format_exc()))
            return False

    @pysnooper.snoop(EasyShellTest().debug)
    def edit(self, newprofile, oldprofile):
        try:
            new = self.sections[self.section_name][newprofile]
            newname = new["Name"]
            newhostname = new['Hostname']
            newusername = new['Username']
            newargument = new['Arguments']
            newlaunchdelay = str(new['Launchdelay'])
            newautolaunch = new['Autolaunch']
            newpersistent = new['Persistent']
            newcustomfile = new['Customfile']
            old = self.sections[self.section_name][oldprofile]
            oldname = old["Name"]
            oldhostname = old['Hostname']
            oldusername = old['Username']
            oldargument = old['Arguments']
            oldautolaunch = old['Autolaunch']
            oldpersistent = old['Persistent']
            oldcustomfile = old['Customfile']
            self.launch()
            EasyshellLib.getElement('KioskMode').Enable()
            EasyshellLib.getElement('DisplayTitle').Enable()
            EasyshellLib.getElement('DisplayConnections').Enable()
            EasyshellLib.getElement('Connections').Click()
            if self.utils(newprofile, 'exist', 'connection'):
                self.utils(newprofile, 'delete', 'connection')
            if self.utils(oldprofile, 'exist', 'connection'):
                self.utils(oldprofile, 'Edit', 'connection')
            else:
                self.Logfile('[Fail]:Old RDP {} do not exist'.format(oldname))
                self.capture('EditRDP', '[Fail]:Old RDP {} do not exist'.format(oldname))
                return False
            time.sleep(5)
            ClearContent(len(oldname) + 5)
            CommonLib.SendKeys(newname)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            ClearContent(len(oldhostname) + 5)
            CommonLib.SendKeys(newhostname)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            ClearContent(len(oldusername) + 5)
            CommonLib.SendKeys(newusername)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
            ClearContent(5)
            CommonLib.SendKeys(str(newlaunchdelay))
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            ClearContent(len(oldargument) + 5)
            CommonLib.SendKeys(newargument)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if newautolaunch == oldautolaunch:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if newpersistent == oldpersistent:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if newcustomfile == oldcustomfile:
                if newcustomfile == 'OFF':
                    CommonLib.SendKey(CommonLib.Keys.VK_TAB)
                else:
                    CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=3)
            else:
                if newcustomfile == 'OFF':
                    CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                    CommonLib.SendKey(CommonLib.Keys.VK_TAB)
                else:
                    if not os.path.exists(newcustomfile):
                        self.Logfile('[Failed]: RDP Connection {} Edit, {} not exist'.format(oldname, newcustomfile))
                        self.capture('EditRDP',
                                     '[Failed]: RDP Connection {} Edit, {} not exist'.format(oldname, newcustomfile))
                        return False
                    CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                    time.sleep(1)
                    CommonLib.SendKey(CommonLib.Keys.VK_TAB)
                    CommonLib.SendKey(CommonLib.Keys.VK_TAB)
                    CommonLib.SendKeys(newcustomfile)
                    CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
            EasyshellLib.getElement('APPLY').Click()
            self.Logfile('[PASS]: RDP Connection {} Edit'.format(oldname))
            return True
        except:
            self.Logfile('[Failed]: RDP Connection {} Edit\n{}'.format(oldprofile, traceback.format_exc()))
            self.capture('EditRDP', '[Failed]: RDP Connection {} Edit\n{}'.format(oldprofile, traceback.format_exc()))
            return False

    @pysnooper.snoop(EasyShellTest().debug)
    def check(self, profile):
        if EasyshellLib.getElement('MAIN_WINDOW').Exists():
            EasyshellLib.getElement('MAIN_WINDOW').SetFocus()
        EasyshellLib.getElement('UserTitles').Click()
        if self.utils(profile, 'exist', 'connection'):
            self.Logfile('[PASS]:RDP connection {} Check Exist'.format(profile))
            return True
        else:
            self.Logfile('[FAIL]:RDP connection {} Check Not Exist'.format(profile))
            self.capture('RDPCheck', '[FAIL]:RDP connection {} Check Not Exist'.format(profile))
            return False


class Shell_Citrix(EasyShellTest):
    def __init__(self):
        EasyShellTest.__init__(self)
        self.section_name = 'createCitrix'

    @pysnooper.snoop(EasyShellTest().debug)
    def create(self, profile):
        test = self.sections[self.section_name][profile]
        name = test["Name"]
        hostname = test['Hostname']
        username = test['Username']
        domain = test['Domain']
        launchdelay = str(test['Launchdelay'])
        autolaunch = test['Autolaunch']
        persistent = test['Persistent']
        try:
            self.launch()
            EasyshellLib.getElement('KioskMode').Enable()
            EasyshellLib.getElement('DisplayTitle').Enable()
            EasyshellLib.getElement('DisplayConnections').Enable()
            EasyshellLib.getElement('Connections').Click()
            if self.utils(profile, 'exist', 'connection'):
                self.utils(profile, 'Delete', 'connection')
            EasyshellLib.getElement('CitrixICAAdd').Click()
            time.sleep(5)
            if EasyshellLib.getElement('CITRIX_EDIT').Exists(3, 3):
                EasyshellLib.getElement('CITRIX_EDIT').SetFocus()
            else:
                self.Logfile('[Fail]: Citrix Edit Window not launched')
                self.capture('CitrixCreate',
                             "[FAIL]: Citrix Edit Window not launched")
                return False
            CommonLib.SendKeys(name)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKeys(hostname)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKeys(str(username))
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKeys(domain)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
            CommonLib.SendKey(CommonLib.Keys.VK_DELETE)
            CommonLib.SendKeys(launchdelay)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if autolaunch == 'OFF' or not autolaunch:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if persistent == 'OFF' or not persistent:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
            EasyshellLib.getElement('APPLY').Click()
            self.Logfile('[PASS]: Citrix Connection {} Create'.format(name))
            return True
        except:
            self.Logfile("[FAIL]: Citrix Connection {} Create\nErrors:\n{}\n".format(name, traceback.format_exc()))
            self.capture('CitrixCreate',
                         "[FAIL]: Citrix Connection {} Create\nErrors:\n{}\n".format(name, traceback.format_exc()))
            return False

    @pysnooper.snoop(EasyShellTest().debug)
    def edit(self, newprofile, oldprofile):
        try:
            new = self.sections[self.section_name][newprofile]
            newname = new["Name"]
            newhostname = new['Hostname']
            newusername = new['Username']
            newdomain = new['Domain']
            newlaunchdelay = str(new['Launchdelay'])
            newautolaunch = new['Autolaunch']
            newpersistent = new['Persistent']
            old = self.sections[self.section_name][oldprofile]
            oldname = old["Name"]
            oldhostname = old['Hostname']
            oldusername = old['Username']
            oldautolaunch = old['Autolaunch']
            oldpersistent = old['Persistent']
            self.launch()
            EasyshellLib.getElement('KioskMode').Enable()
            EasyshellLib.getElement('DisplayTitle').Enable()
            EasyshellLib.getElement('DisplayConnections').Enable()
            EasyshellLib.getElement('Connections').Click()
            if self.utils(newprofile, 'exist', 'connection'):
                self.utils(newprofile, 'delete', 'connection')
            if self.utils(oldprofile, 'exist', 'connection'):
                self.utils(oldprofile, 'Edit', 'connection')
            else:
                self.Logfile('[Fail]:Old CitrixICA {} do not exist'.format(oldname))
                self.capture('CitrixEdit', '[Fail]:Old CitrixICA {} do not exist'.format(oldname))
                return False
            time.sleep(5)
            if EasyshellLib.getElement('CITRIX_EDIT').Exists(3, 3):
                EasyshellLib.getElement('CITRIX_EDIT').SetFocus()
            else:
                self.Logfile('[Fail]: Citrix Edit Window not launched')
                self.capture('CitrixCreate',
                             "[FAIL]: Citrix Edit Window not launched")
                return False
            ClearContent(len(oldname) + 5)
            CommonLib.SendKeys(newname)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            ClearContent(len(oldhostname) + 5)
            CommonLib.SendKeys(newhostname)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            ClearContent(len(oldusername) + 5)
            CommonLib.SendKeys(newusername)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            ClearContent(10)
            CommonLib.SendKeys(newdomain)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            ClearContent(5)
            CommonLib.SendKeys(newlaunchdelay)
            CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if newautolaunch == oldautolaunch:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if newpersistent == oldpersistent:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
            EasyshellLib.getElement('APPLY').Click()
            self.Logfile('[PASS]: Citrix Connection {} Edit'.format(oldname))
            return True
        except:
            self.Logfile('[Failed]: Citrix Connection {} Edit\n{}'.format(oldprofile, traceback.format_exc()))
            self.capture('CitrixEdit',
                         '[Failed]: Citrix Connection {} Edit\n{}'.format(oldprofile, traceback.format_exc()))
            return False

    @pysnooper.snoop(EasyShellTest().debug)
    def check(self, profile):
        if EasyshellLib.getElement('MAIN_WINDOW').Exists():
            EasyshellLib.getElement('MAIN_WINDOW').SetFocus()
        EasyshellLib.getElement('UserTitles').Click()
        if self.utils(profile, 'exist', 'connection'):
            self.Logfile('[PASS]:CitrixICA connection {} Check Exist'.format(profile))
            return True
        else:
            self.Logfile('[FAIL]:CitrixICA connection {} Check Not Exist'.format(profile))
            self.capture('CitrixCheck', '[FAIL]:CitrixICA connection {} Check Not Exist'.format(profile))
            return False


class TaskSwitcher(EasyShellTest):
    def __init__(self):
        EasyShellTest.__init__(self)

    def __prepare(self):
        self.launch()
        EasyshellLib.getElement('Settings').Click()
        EasyshellLib.getElement('KioskMode').Enable()

    def enable(self):
        try:
            self.__prepare()
            EasyshellLib.getElement('EnableTaskSwitcher').Enable()
            EasyshellLib.getElement('APPLY').Click()
            self.Logfile("[PASS]: enable Task Switcher")
            return True
        except:
            self.Logfile("[Fail]: enable Task Switcher")
            self.capture("[Fail]: enable Task Switcher")
            return False

    def disable(self):
        try:
            self.__prepare()
            EasyshellLib.getElement('EnableTaskSwitcher').Disable()
            EasyshellLib.getElement('APPLY').Click()
            self.Logfile("[PASS]: disable Task Switcher")
            return True
        except:
            self.Logfile("[Fail]: Disable Task Switcher")
            self.capture("[Fail]: Disable Task Switcher")
            return False

    def enablePermanent(self):
        try:
            self.enable()
            EasyshellLib.getElement('Permanent').Enable()
            EasyshellLib.getElement('APPLY').Click()
            return True
        except:
            self.Logfile("[Fail]: enable Task Switcher")
            self.capture("[Fail]: enable Task Switcher")
            return False

    def disablePermanent(self):
        try:
            self.enable()
            EasyshellLib.getElement('Permanent').Disable()
            EasyshellLib.getElement('APPLY').Click()
            return True
        except:
            self.Logfile('[Fail]: Enable permanently\n {}'.format(traceback.format_exc()))
            self.capture('[Fail]: Enable permanently\n {}'.format(traceback.format_exc()))
            return False

    def checkPermanent(self):
        if not EasyshellLib.getElement("TASK_SWITCHER").Exists(1, 1):
            self.Logfile("[Fail]: Task SwitcherBar is not shown")
            self.capture("[Fail]: Task switcherBar is not Shown")
            return False
        time.sleep(15)
        if EasyshellLib.getElement("TASK_SWITCHER").BoundingRectangle[0] == 0:
            self.Logfile("[PASS]: Task SwitcherBar is permanent")
            return True
        else:
            self.capture("[Fail]: Task SwitcherBar is not permanent, expect permanent")
            self.Logfile("[Fail]: Task SwitcherBar is not permanent, expect permanent")
            return False

    def checkNoPermanent(self):
        if not EasyshellLib.getElement("TASK_SWITCHER").Exists(1, 1):
            self.Logfile("[Fail]: Task SwitcherBar is shown, expect No permanent")
            self.capture("[Fail]: Task SwitcherBar is shown, expect No permanent")
            return False
        time.sleep(15)
        if EasyshellLib.getElement("TASK_SWITCHER").BoundingRectangle[0] == 0:
            self.Logfile("[Fail]: Task SwitcherBar is permanent, expect no permanent")
            self.capture("[Fail]: Task SwitcherBar is permanent, expect no permanent")
            return False
        else:
            self.Logfile("[PASS]: Task SwitcherBar is not permanent")
            return True

    def enableSoundIconReadOnly(self):
        self.enable()
        EasyshellLib.getElement('DisplaySoundIconInteraction').Disable()
        EasyshellLib.getElement('Permanent').Enable()
        EasyshellLib.getElement('APPLY').Click()
        EasyshellLib.getElement('Exit').Click()
        self.Logfile("[PASS]: sound icon read only settings")

    def checkSoundReadOnly(self):
        EasyshellLib.getElement('SoundIcon').Click()
        if EasyshellLib.getElement('SoundAdjust').IsOffScreen:
            self.Logfile("[PASS]: sound value is not shown")
            return True
        else:
            self.Logfile("[Fail]: sound value is shown, Expect not shown")
            self.capture('SoundReadOnly', "[Fail]: sound value is shown, expect not shown")
            return False

    def enableSoundInteraction(self):
        self.enable()
        EasyshellLib.getElement('DisplaySound').Enable()
        EasyshellLib.getElement('Permanent').Enable()
        EasyshellLib.getElement('DisplaySoundIconInteraction').Enable()
        EasyshellLib.getElement('APPLY').Click()
        self.Logfile("[PASS]: enable sound interaction")

    def checkSoundInteraction(self):
        """
        判断sound interaction 是否工作
        1. 判断当前音量大小，大于60时降低音量操作，小于80时增加音量操作
        2. 通过鼠标键盘两种操作来测试功能
        """
        EasyshellLib.getElement('SoundIcon').Click()
        time.sleep(3)
        if not EasyshellLib.getElement('SoundAdjust').IsOffScreen:
            # EasyshellLib.getElement('SoundIcon').Click()
            currentVol = EasyshellLib.getElement('SoundAdjust').AccessibleCurrentValue()
            currentThumb = EasyshellLib.getElement('SoundAdjustBar').BoundingRectangle
            if int(currentVol) < 60:
                # increase volumn +10
                EasyshellLib.CommonLib.DragDrop(currentThumb[0], currentThumb[1], currentThumb[0]+10, currentThumb[1])
                tempVol = EasyshellLib.getElement('SoundAdjust').AccessibleCurrentValue()
                if currentVol != tempVol:
                    self.Logfile("[PASS]: Sound Adjusted by Mouse")
                else:
                    self.Logfile("[FAIL]: Sound Adjusted by Mouse")
                    return False
                CommonLib.SendKey(CommonLib.Keys.VK_RIGHT)
                finalVol = EasyshellLib.getElement('SoundAdjust').AccessibleCurrentValue()
                if finalVol != tempVol:
                    self.Logfile("[PASS]: Sound Adjust by Keyboard")
                    return True
                else:
                    self.Logfile("[FAIL]: Sound Adjust by Keyboard")
                    return False
            else:
                EasyshellLib.getElement('SoundAdjustBar').Drag(-10, 0)
                tempVol = EasyshellLib.getElement('SoundAdjust').AccessibleCurrentValue()
                if currentVol != tempVol:
                    self.Logfile("[PASS]: Sound Adjust by Mouse")
                else:
                    self.Logfile("[FAIL]: Sound Adjust by Mouse")
                    return False
                CommonLib.SendKey(CommonLib.Keys.VK_LEFT)
                finalVol = EasyshellLib.getElement('SoundAdjust').AccessibleCurrentValue()
                if finalVol != tempVol:
                    self.Logfile("[PASS]: {} Sound Adjust by Keyboard")
                    return True
                else:
                    self.Logfile("[FAIL]: {} Sound Adjust by Keyboard")
                    return False
        else:
            self.Logfile("[Fail]: sound value is not shown")
            return False


class General_Test(EasyShellTest):
    def __init__(self):
        EasyShellTest.__init__(self)

    @staticmethod
    def create_user(user='standard_user', password='test'):
        user = CommonLib.User_Group(user, password, 'Users')
        user.add_user()
        user.add_user_to_group()

    @staticmethod
    def create_admin(user='standard_admin', passwd='test'):
        user = CommonLib.User_Group(user, passwd, 'Administrators')
        user.add_user()
        user.add_user_to_group()

    @staticmethod
    def enable_kiosk():
        reg = CommonLib.Reg_Utils()
        key = reg.isKeyExist(r'SOFTWARE\HP\HP Easy Shell')
        if key:
            reg.create_value(key, 'KioskMode', 0, 'True')
            reg.close(key)

    def reg_export(self):
        self.launch()
        EasyshellLib.getElement('Export').Click()
        if not os.path.exists(r'c:\temp'):
            os.mkdir(r'c:\temp')
        EasyshellLib.getElement('SaveToFile').SetValue(r'c:\temp\easyshellsettings.reg')
        EasyshellLib.getElement('ButtonSave').Click()
        time.sleep(1)
        if EasyshellLib.getElement('OVERRIDE').Exists():
            EasyshellLib.getElement('ButtonYES').Click()
        EasyshellLib.getElement('Exit').Click()
        if os.path.exists(r'c:\temp\easyshellsettings.reg'):
            self.Logfile('[PASS]:export settings to c:/temp/easyshellsettings.reg')
            return True
        else:
            self.Logfile('[FAIL]:export function fail, c:/temp/easyshellsettings.reg not exist')
            self.capture('ExportSettings', '[FAIL]:export function fail, c:/temp/easyshellsettings.reg not exist')
            return False

    def reg_import(self):
        os.popen(r'c:/temp/easyshellsettings.reg')
        time.sleep(2)
        if EasyshellLib.getElement('REG_EDIT_WARNNING').Exists(0, 0):
            EasyshellLib.getElement('ButtonYES').Click()
            time.sleep(2)
            if EasyshellLib.getElement('REG_EDIT').Exists(0, 0):
                EasyshellLib.getElement('ButtonOK').Click()
                self.Logfile('[PASS]:c:/temp/easyshellsettings.reg import successfully')
        else:
            self.Logfile('[FAIL]:c:/temp/easyshellsettings.reg import FAIL, please double check reg file')
            self.capture('ImportSettings', 'c:/temp/easyshellsettings.reg import FAIL, please double check reg file')
            return False

    def modify(self):
        self.launch()
        EasyshellLib.getElement('Applications').Click()
        self.utils('resetapp', 'delete')
        EasyshellLib.getElement('StoreFront').Click()
        self.utils('resetapp', 'delete')
        EasyshellLib.getElement('WebSites').Click()
        self.utils('resetapp', 'delete')

    def import_check(self, imported=True):
        EasyshellLib.getElement("UserTitles").Click()
        app = CommonLib.TextControl(Name='test_app').Exists()
        store = CommonLib.TextControl(Name='test_store').Exists()
        web = CommonLib.TextControl(Name='test_web').Exists()
        if imported:
            if app and store and web:
                self.Logfile('[PASS]: check import setting PASS')
                return True
            else:
                self.Logfile('[FAIL]: check import setting Fail:\napplication exist:{}\n'
                             'stoer exist:{}\nweb exist:{}'.format(app, store, web))
                self.capture('import_check', '[FAIL]: check setting Fail:\napplication exist:{}\n'
                                             'stoer exist:{}\nweb exist:{}'.format(app, store, web))
                return False
        else:
            if app or store or web:
                self.Logfile('[FAIL]: check imoprt setting Fail(before import):\napplication exist:{}\n'
                             'stoer exist:{}\nweb exist:{}'.format(app, store, web))
                self.capture('import_check', '[FAIL]: check setting Fail(before import):\napplication exist:{}\n'
                                             'stoer exist:{}\nweb exist:{}'.format(app, store, web))
                return False

            else:
                self.Logfile('[PASS]: check import setting PASS(before import)')
                return True

    def __exit_menu(self):
        EasyshellLib.getElement('AdminPower').Click()
        EasyshellLib.getElement('ExitItem').Click()
        time.sleep(1)
        if EasyshellLib.getElement('ExitSave').Exists():
            self.Logfile('[PASS]: Exit Menu Test Pass, save dialog pops up')
            return True
        else:
            self.Logfile('[FAIL]: Exist Menu Test Fail, save dialog not pop up')
            self.capture('Exit_Menu', '[FAIL]: Exist Menu Test Fail, save dialog not pop up')
            return False

    def __exit_button(self):
        EasyshellLib.getElement('Exit').Click()
        time.sleep(1)
        if EasyshellLib.getElement('ExitSave').Exists():
            self.Logfile('[PASS]: Exit Button Test Pass, save dialog pops up')
            return True
        else:
            self.Logfile('[FAIL]: Exist Button Test Fail, save dialog not pop up')
            self.capture('Exit_Button', '[FAIL]: Exist Button Test Fail, save dialog not pop up')
            return False

    def __exit_check(self, item, button, state):
        if button.upper() == 'YES':
            EasyshellLib.getElement('ButtonYES').Click()
            self.launch()
            if EasyshellLib.getElement('AllowLock').GetStatus() != state:
                EasyshellLib.getElement('Exit').Click()
                self.Logfile('[PASS]: Exist {} Test PASS'.format(item))
                return True
            else:
                EasyshellLib.getElement('Exit').Click()
                self.Logfile('[FAIL]: Exit {} Test Fail, Expect AllowLock button {}'
                             ''.format(item, 'Enable' if state == 1 else 'Disabled'))
                self.capture('Exit_{}'.format(item), '[FAIL]: Exit {} Test Fail, Expect AllowLock button {}'
                                                     ''.format(item, 'Enable' if state == 1 else 'Disabled'))
                return False
        elif button.upper() == 'NO':
            EasyshellLib.getElement('ButtonNO').Click()
            self.launch()
            if EasyshellLib.getElement('AllowLock').GetStatus() == state:
                EasyshellLib.getElement('Exit').Click()
                self.Logfile('[PASS]: Exit {} Test PASS'.format(item))
                return True
            else:
                EasyshellLib.getElement('Exit').Click()
                self.Logfile('[FAIL]: Exit {} Test Fail, Expect AllowLock button {}'
                             ''.format(item, 'Enable' if state == 1 else 'Disabled'))
                self.capture('Exit_{}'.format(item), '[FAIL]: Exit {} Test Fail, Expect AllowLock button {}'
                                                     ''.format(item, 'Enable' if state == 1 else 'Disabled'))
                return False
        else:
            print('aaa')
            pass

    def exit_button(self, button):
        self.launch()
        state = EasyshellLib.getElement('AllowLock').GetStatus()
        print(type(state))
        if state:
            EasyshellLib.getElement('AllowLock').Disable()
        else:
            EasyshellLib.getElement('AllowLock').Enable()
        # ---------test exit button -------------------------------------------
        if self.__exit_button():
            return self.__exit_check('button', button, state)
        else:
            return False

    def exit_menu(self, button):
        self.launch()
        state = EasyshellLib.getElement('AllowLock').GetStatus()
        if EasyshellLib.getElement('AllowLock').GetStatus():
            EasyshellLib.getElement('AllowLock').Disable()
        else:
            EasyshellLib.getElement('AllowLock').Enable()
        # ---------test exit menu -------------------------------------------
        if self.__exit_menu():
            return self.__exit_check('menu', button, state)
        else:
            return False

    def set_hide_during_session(self, state):
        General_Test().resetEasyshell()
        self.launch()
        EasyshellLib.getElement('Settings').Click()
        if state.upper() == 'ON':
            EasyshellLib.getElement('HideEasyShell').Enable()
        else:
            EasyshellLib.getElement('HideEasyShell').Disable()
        EasyshellLib.getElement('APPLY').Click()

    def check_hide_during_session(self, state):
        EasyshellLib.CommonLib.TextControl(Name='test_app').Click()
        if state:
            if not EasyshellLib.getElement('MAIN_WINDOW').Exists():
                print('pass')
                EasyshellLib.CommonLib.WindowControl(RegexName='.*Notepad').Close()
                return True
            else:
                print('fail')
                return False
        else:
            if EasyshellLib.getElement('MAIN_WINDOW').Exists():
                EasyshellLib.CommonLib.WindowControl(RegexName='.*Notepad').Close()
                print('pass')
                return True
            else:
                print('fail')
                return False

    def hex2rgb(self, hex_str):
        real_str = hex_str[-6:]
        splited = re.findall(r'(.{2})', real_str)
        return tuple([int(i, 16) for i in splited])

    def check_copyright(self):
        self.launch()
        flag = True
        version = EasyshellLib.getElement('CopyRight').Name
        ver = re.findall(r'(.*)Copy.*', version)[0].replace('-', '').strip()
        copy = re.findall(r'.* Copyright(.*)HP', version)[0][3:].strip()
        comp = re.findall(r'.*-\d\d\d\d(.*)', version)[0].strip()
        if ver != self.sections['version']:
            flag = False
            self.Logfile('[FAIL]:Version Check Fail, expect:{}'.format(self.sections['version']))
            self.capture('check_version', '[FAIL]:Version Check Fail, expect:{}'.format(self.sections['version']))
        if copy != self.sections['copyright']:
            flag = False
            self.Logfile('[FAIL]:copyright Check Fail, expect:{}'.format(self.sections['copyright']))
            self.capture('check_copyright', '[FAIL]:copyright Check Fail, expect:{}'.format(self.sections['copyright']))
        if comp != self.sections['company']:
            flag = False
            self.Logfile('[FAIL]:company Check Fail, expect:{}'.format(self.sections['company']))
            self.capture('check_company', '[FAIL]:company Check Fail, expect:{}'.format(self.sections['company']))
        EasyshellLib.getElement('Exit').Click()
        return flag


if __name__ == '__main__':
    # import uiautomation
    # rect = uiautomation.ThumbControl(AutomationId='Thumb').BoundingRectangle
    # uiautomation.DragDrop(rect[0], rect[1], rect[0]-10, rect[1])
    # TaskSwitcher().enableSoundInteraction()
    TaskSwitcher().checkSoundInteraction()
    pass

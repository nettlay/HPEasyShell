import Test_Scripts.EasyShell_Lib as EasyshellLib
import Library.CommonLib as CommonLib
import os
import time
import traceback
import pysnooper


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
        self.section_name = ''
        #  ---------------------------------
        self.log_path = os.path.join(self.path, 'Test_Report')
        self.misc = os.path.join(self.path, 'Misc')
        self.data = os.path.join(self.path, 'Test_Data')
        self.casepath = os.path.join(self.path, 'Test_Suite')
        self.testset = os.path.join(self.path, 'testset.xlsx')
        self.sections = CommonLib.YmlUtils(os.path.join(self.data, "easyshell_testdata.yaml")).get_item()
        self.appPath = self.sections['appPath']['easyShellPath']
        self.debug = os.path.join(self.log_path, 'debug.log')

    def create(self, profile):
        pass

    def edit(self, newprofile, oldprofile):
        pass

    def check(self, profile):
        pass

    # ------------------------------ Utils -------------------------------------------
    def Logfile(self, rs):
        EasyshellLib.TxtUtils(os.path.join(self.log_path, "easyShellLog.txt"), 'a').set_msg(
            "[{}]:{}\n".format(time.ctime(), rs))

    def utils(self, profile='', op='exist', item='normal'):
        """
        :param profile:  test profile, [test1,test2,,standardApp...]
        :param op: test option [exist | notexist | edit | delete |launch]
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
        edit = appControl.ButtonControl(AutomationId='editButton')
        delete = appControl.ButtonControl(AutomationId='deleteButton')
        if op.upper() == 'LAUNCH':
            launch.Click()
            return True
        elif op.upper() == 'EDIT':
            edit.Click()
            return True
        elif op.upper() == 'DELETE':
            try:
                delete.Click()
                EasyshellLib.getElement('DeleteYes').Click()
                EasyshellLib.getElement('APPLY').Click()
                return True
            except:
                self.Logfile("[FAIL]:App {} Delete\nErrors:\n{}\n".format(name, traceback.format_exc()))
                return False
        else:
            return True


class UserInterfacSettings(EasyShellTest):
    def __init__(self):
        EasyShellTest.__init__(self)

    def modify(self, profile):
        """
        :param profile: one of test parameters' combination
        """
        flag = True  # record the function's status
        content = self.sections
        test = content['userInterface'][profile]
        for app_path in self.appPath:
            # launch app from file according given file path
            if os.path.exists(app_path):
                EasyshellLib.CommonUtils.LaunchAppFromFile(app_path)
                break
            else:
                continue
        time.sleep(5)
        self.Logfile('---------Begin To Test Modify settings----------')
        EasyshellLib.getElement('KioskMode').Enable()
        for item in test:
            name = item.split(":")[0].strip()  # setting name
            status = item.split(":")[1].strip()  # setting status on/off
            if status == 'ON':
                try:
                    EasyshellLib.getElement(name).Enable()
                except:
                    flag = False
                    self.Logfile('[Fail]Button {} Enable\n{}'.format(name, traceback.format_exc()))
            elif status == 'OFF':
                try:
                    EasyshellLib.getElement(name).Disable()
                except:
                    flag = False
                    self.Logfile('[Fail]Button {} Disable\n{}'.format(name, traceback.format_exc()))
            else:
                flag = False
                self.Logfile('[Fail]Button {} Status in test data is not Correct!'.format(name))
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
        swTitles = True
        swBrowser = True
        swPower = True
        self.Logfile('---------Begin To Test Check Modify settings----------')
        for item in test:
            """
            1. 没有测试总开关关闭但是子开关开启时的状态
            2. Wifi 只测试链接窗口是否弹出，没有测试wifi连接功能
            """
            name = item.split(":")[0].strip()
            status = item.split(":")[1].strip()
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
                    if not EasyshellLib.getElement('UserTitles').IsOffscreen:
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'DisplayBrowser':
                    if not EasyshellLib.getElement('UserBrowser').IsOffscreen:
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'DisplayAdmin':
                    if not EasyshellLib.getElement('UserAdmin').IsOffscreen:
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'DisplayPower':
                    if not EasyshellLib.getElement('UserPower').IsOffscreen:
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                # ////////// item for Titles //////////////////////////////////////
                if name == 'DisplayApp':
                    EasyshellLib.getElement('UserTitles').Click()
                    time.sleep(1)
                    if not EasyshellLib.getElement('UserApp').IsOffscreen:
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'DisplayConnections':
                    EasyshellLib.getElement('UserTitles').Click()
                    time.sleep(1)
                    if not EasyshellLib.getElement('UserConnection').IsOffscreen:
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'DisplayStoreFront':
                    EasyshellLib.getElement('UserTitles').Click()
                    time.sleep(1)
                    if not EasyshellLib.getElement('UserStoreFront').IsOffscreen():
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'DisplayWebsites':
                    EasyshellLib.getElement('UserTitles').Click()
                    time.sleep(1)
                    if not EasyshellLib.getElement('UserWebsites').IsOffscreen:
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                # ////////item for Web browser ///////////////////////////////
                if name == 'DisplayAddress':
                    EasyshellLib.getElement('UserBrowser').Click()
                    time.sleep(1)
                    if not EasyshellLib.getElement('AddressBar').IsOffscreen:
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'DisplayHome':
                    EasyshellLib.getElement('UserBrowser').Click()
                    time.sleep(1)
                    if not EasyshellLib.getElement('WebHome').IsOffscreen:
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                # ------------------------Item for Admin power----------------------------------------
                if name == 'AllowLock':
                    EasyshellLib.getElement('UserPower').Click()
                    time.sleep(1)
                    if not EasyshellLib.getElement('Lock').IsOffscreen:
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'AllowLogoff':
                    EasyshellLib.getElement('UserPower').Click()
                    time.sleep(1)
                    if not EasyshellLib.getElement('Logoff').IsOffscreen:
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'AllowRestart':
                    EasyshellLib.getElement('UserPower').Click()
                    time.sleep(1)
                    if not EasyshellLib.getElement('Restart').IsOffscreen:
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'AllowShutDown':
                    EasyshellLib.getElement('UserPower').Click()
                    time.sleep(1)
                    if not EasyshellLib.getElement('Shutdown').IsOffscreen:
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                # ----------------- Virtual keyboard --------
                if name == 'DisplayVKeyboard':
                    if not EasyshellLib.getElement('UserKeyBoard').IsOffscreen:
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                # ---Label that display mac/time/version... at the bottom of UI -----
                if name == 'DisplayTime':
                    if not EasyshellLib.getElement('Time').IsOffscreen:
                        self.Logfile("[PASS]: {} is shown".format(name))
                        real_time = EasyshellLib.CommonUtils.getLocalTime('%H:%M')
                        show_time = EasyshellLib.getElement('Time').Name
                        if real_time.split(':')[0] in show_time:
                            self.Logfile("-->[PASS]: {} real time format".format(name))
                        else:
                            self.Logfile("-->[Fail]: {} real time format".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'DisplayIP':
                    if not EasyshellLib.getElement('IPAddr').IsOffscreen:
                        self.Logfile("[PASS]: {} is shown".format(name))
                        print(EasyshellLib.CommonUtils.getNetInfo(), '--------net info')
                        real_ip = EasyshellLib.CommonUtils.getNetInfo()['IP']
                        show_ip = EasyshellLib.getElement('IPAddr').Name
                        if real_ip == show_ip:
                            self.Logfile("-->[PASS]: {} real IP".format(name))
                        else:
                            self.Logfile("-->[Fail]: {} real IP".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'DisplayMAC':
                    if not EasyshellLib.getElement('MACAddr').IsOffscreen:
                        self.Logfile("[PASS]: {} is shown".format(name))
                        real_mac = EasyshellLib.CommonUtils.getNetInfo()['MAC']
                        show_mac = EasyshellLib.getElement('MACAddr').Name
                        if real_mac == show_mac:
                            self.Logfile("-->[PASS]: {} real mac".format(name))
                        else:
                            self.Logfile("-->[Fail]: {} real mac".format(name))

                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
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
                    if EasyshellLib.getElement('UserTitles').IsOffscreen:
                        swTitles = False
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'DisplayBrowser':
                    if EasyshellLib.getElement('UserBrowser').IsOffscreen():
                        swBrowser = False
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'DisplayAdmin':
                    if EasyshellLib.getElement('UserAdmin').IsOffscreen:
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'DisplayPower':
                    if EasyshellLib.getElement('UserPower').IsOffscreen:
                        swPower = False
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                # ////////// item for Titles //////////////////////////////////////
                if name == 'DisplayApp':
                    EasyshellLib.getElement('UserTitles').Click()
                    time.sleep(1)
                    if EasyshellLib.getElement('UserApp').IsOffscreen:
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'DisplayConnections':
                    EasyshellLib.getElement('UserTitles').Click()
                    time.sleep(1)
                    if EasyshellLib.getElement('UserConnection').IsOffscreen:
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'DisplayStoreFront':
                    EasyshellLib.getElement('UserTitles').Click()
                    time.sleep(1)
                    if EasyshellLib.getElement('UserStoreFront').IsOffscreen:
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'DisplayWebsites':
                    EasyshellLib.getElement('UserTitles').Click()
                    time.sleep(1)
                    if EasyshellLib.getElement('UserBrowser').IsOffscreen:
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                # ////////item for Web browser ///////////////////////////////
                if name == 'DisplayAddress':
                    EasyshellLib.getElement('UserBrowser').Click()
                    time.sleep(1)
                    if EasyshellLib.getElement('AddressBar').IsOffscreen:
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'DisplayHome':
                    EasyshellLib.getElement('UserBrowser').Click()
                    time.sleep(1)
                    if EasyshellLib.getElement('WebHome').IsOffscreen:
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                # ------------------------Item for Admin power----------------------------------------
                if name == 'AllowLock':
                    EasyshellLib.getElement('UserPower').Click()
                    time.sleep(1)
                    if EasyshellLib.getElement('Lock').IsOffscreen:
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'AllowLogoff':
                    EasyshellLib.getElement('UserPower').Click()
                    time.sleep(1)
                    if EasyshellLib.getElement('Logoff').IsOffscreen:
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'AllowRestart':
                    EasyshellLib.getElement('UserPower').Click()
                    time.sleep(1)
                    if EasyshellLib.getElement('Restart').IsOffscreen:
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'AllowShutDown':
                    EasyshellLib.getElement('UserPower').Click()
                    time.sleep(1)
                    if EasyshellLib.getElement('Shutdown').IsOffscreen():
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                # ----------------- Virtual keyboard --------
                if name == 'DisplayVKeyboard':
                    if EasyshellLib.getElement('UserKeyBoard').IsOffscreen:
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                # ---Label that display mac/time/version... at the bottom of UI -----
                if name == 'DisplayTime':
                    if EasyshellLib.getElement('Time').IsOffscreen:
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'DisplayIP':
                    if EasyshellLib.getElement('IPAddr').IsOffscreen:
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'DisplayMAC':
                    if EasyshellLib.getElement('MACAddr').IsOffscreen:
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                # ------------No Enable Network status notification -----------
                # ------------No Hide HP Easy Shell during Session ------------
        return flag


class UserSettings(EasyShellTest):
    def __init__(self):
        EasyShellTest.__init__(self)

    def modify(self, profile):
        """
        :param profile: one of test parameters' combination
        """
        flag = True  # record the function's status
        content = self.sections
        test = content['userSettings'][profile]
        for app_path in self.appPath:
            # launch app from file according given file path
            if os.path.exists(app_path):
                EasyshellLib.CommonUtils.LaunchAppFromFile(app_path)
                break
            else:
                continue
        time.sleep(5)
        self.Logfile('---------Begin To Test Modify User settings----------')
        EasyshellLib.getElement('KioskMode').Enable()
        for item in test:
            name = item.split(":")[0].strip()  # setting name
            status = item.split(":")[1].strip()  # setting status on/off
            if status == 'ON':
                try:
                    EasyshellLib.getElement(name).Enable()
                except:
                    flag = False
                    self.Logfile('[Fail]Button {} Enable\n{}'.format(name, traceback.format_exc()))
            elif status == 'OFF':
                try:
                    EasyshellLib.getElement(name).Disable()
                except:
                    flag = False
                    self.Logfile('[Fail]Button {} Disable\n{}'.format(name, traceback.format_exc()))
            else:
                flag = False
                self.Logfile('[Fail]Button {} Status in test data is not Correct!'.format(name))
        EasyshellLib.getElement('APPLY').Click()
        EasyshellLib.getElement('Exit').Click()
        self.Logfile('[PASS] Modify user settings')
        return flag

    def check(self, profile):
        flag = True
        # ---Below test use yml, list type can be test in order --------
        content = self.sections
        test = content['userInterface'][profile]
        # 以下为部分设置的总开关sw = switch
        self.Logfile('---------Begin To Test Check User settings----------')
        swSettings = True
        for item in test:
            """
            1. 没有测试总开关关闭但是子开关开启时的状态
            2. Wifi 只测试链接窗口是否弹出，没有测试wifi连接功能
            """
            name = item.split(":")[0].strip()
            status = item.split(":")[1].strip()
            if status == "ON":
                # ///////判断主按钮是否关闭，如果OFF,子按钮不再检查
                if not swSettings:
                    if name in ['AllowMouse', 'AllowKeyboard', 'AllowDisplay', 'AllowSound', 'AllowRegion',
                                'AllowNetworkConn', 'AllowDateTime', 'AllowEasyAccess', 'AllowIEProperty',
                                'AllowWifiConfig']:
                        continue
                # ------------User Settings -----------------------------------
                if name == 'AllowMouse':
                    EasyshellLib.getElement('UserSettings').Click()
                    time.sleep(1)
                    if not EasyshellLib.getElement('SysMouseIcon').IsOffscreen:
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'AllowKeyboard':
                    EasyshellLib.getElement('UserSettings').Click()
                    time.sleep(1)
                    if not EasyshellLib.getElement('SysKeyboardIcon').IsOffscreen:
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'AllowDisplay':
                    EasyshellLib.getElement('UserSettings').Click()
                    time.sleep(1)
                    if not EasyshellLib.getElement('SysDisplayIcon').IsOffscreen:
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'AllowSound':
                    EasyshellLib.getElement('UserSettings').Click()
                    time.sleep(1)
                    if not EasyshellLib.getElement('SysSoundIcon').IsOffscreen:
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'AllowRegion':
                    EasyshellLib.getElement('UserSettings').Click()
                    time.sleep(1)
                    if not EasyshellLib.getElement('SysRegionIcon').IsOffscreen:
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'AllowNetworkConn':
                    EasyshellLib.getElement('UserSettings').Click()
                    time.sleep(1)
                    if not EasyshellLib.getElement('SysNetworkConnIcon').IsOffscreen:
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'AllowDateTime':
                    EasyshellLib.getElement('UserSettings').Click()
                    time.sleep(1)
                    if not EasyshellLib.getElement('SysDateTimeIcon').IsOffscreen:
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'AllowEasyAccess':
                    EasyshellLib.getElement('UserSettings').Click()
                    time.sleep(1)
                    if not EasyshellLib.getElement('SysEaseAccessCenterIcon').IsOffscreen:
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'AllowIEProperty':
                    EasyshellLib.getElement('UserSettings').Click()
                    time.sleep(1)
                    if not EasyshellLib.getElement('SysIEIcon').IsOffscreen:
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'AllowWifiConfig':
                    EasyshellLib.getElement('UserSettings').Click()
                    time.sleep(1)
                    if not EasyshellLib.getElement('SysWirelessIcon').IsOffscreen:
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
            elif status == "OFF":
                # ///////判断主按钮是否关闭，如果OFF,子按钮不再检查
                if not swSettings:
                    if name in ['AllowMouse', 'AllowKeyboard', 'AllowDisplay', 'AllowSound', 'AllowRegion',
                                'AllowNetworkConn', 'AllowDateTime', 'AllowEasyAccess', 'AllowIEProperty',
                                'AllowWifiConfig']:
                        continue
                # ------------User Settings -----------------------------------
                if name == 'AllowMouse':
                    EasyshellLib.getElement('UserSettings').Click()
                    time.sleep(1)
                    if not EasyshellLib.getElement('SysMouseIcon').Exists(0, 0):
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'AllowKeyboard':
                    EasyshellLib.getElement('UserSettings').Click()
                    time.sleep(1)
                    if not EasyshellLib.getElement('SysKeyboardIcon').Exists(0, 0):
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'AllowDisplay':
                    EasyshellLib.getElement('UserSettings').Click()
                    time.sleep(1)
                    if not EasyshellLib.getElement('SysDisplayIcon').Exists(0, 0):
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'AllowSound':
                    EasyshellLib.getElement('UserSettings').Click()
                    time.sleep(1)
                    if not EasyshellLib.getElement('SysSoundIcon').Exists(0, 0):
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'AllowRegion':
                    EasyshellLib.getElement('UserSettings').Click()
                    time.sleep(1)
                    if not EasyshellLib.getElement('SysRegionIcon').Exists(0, 0):
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'AllowNetworkConn':
                    EasyshellLib.getElement('UserSettings').Click()
                    time.sleep(1)
                    if not EasyshellLib.getElement('SysNetworkConnIcon').Exists(0, 0):
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'AllowDateTime':
                    EasyshellLib.getElement('UserSettings').Click()
                    time.sleep(1)
                    if not EasyshellLib.getElement('SysDateTimeIcon').Exists(0, 0):
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'AllowEasyAccess':
                    EasyshellLib.getElement('UserSettings').Click()
                    time.sleep(1)
                    if not EasyshellLib.getElement('SysEaseAccessCenterIcon').Exists(0, 0):
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'AllowIEProperty':
                    EasyshellLib.getElement('UserSettings').Click()
                    time.sleep(1)
                    if not EasyshellLib.getElement('SysIEIcon').Exists(0, 0):
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'AllowWifiConfig':
                    EasyshellLib.getElement('UserSettings').Click()
                    time.sleep(1)
                    if not EasyshellLib.getElement('SysWirelessIcon').Exists(0, 0):
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
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
        for app_path in self.appPath:
            if os.path.exists(app_path):
                print(app_path)
                EasyshellLib.CommonUtils.LaunchAppFromFile(self.appPath)
                break
            else:
                continue
        time.sleep(5)
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
        content = CommonLib.YmlUtils(os.path.join(self.data, "easyshell_testdata.yaml")).get_item()
        test = content['createApp'][profile]
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
            if AdminOnly:
                if self.utils(profile, 'exist'):
                    flag = False
                    self.Logfile("[Fail]:App {} AdminOnly: {}".format(Name, AdminOnly))
                    return flag
                else:
                    flag = True
                    self.Logfile("[PASS]:App {} AdminOnly:{}".format(Name, AdminOnly))
                    return flag
            else:
                if not self.utils(profile, 'exist'):
                    flag = False
                    self.Logfile("[Fail]:App {} AdminOnly:{}".format(Name, AdminOnly))
                    return flag
            if HideMissApp:
                if self.utils(profile, 'exist'):
                    flag = False
                    self.Logfile("[Failed]:App {} Hide Missing App".format(Name))
                    return flag
                else:
                    self.Logfile("[PASS]:App {} Hide Missing App".format(Name))
                    return flag
            if Autolaunch == 0:
                self.utils(profile, "launch")
                if Launchdelay == 0:
                    for t in range(5):
                        if CommonLib.WindowControl(RegexName=WindowName).Exists(0, 0):
                            break
                        else:
                            continue
                    if not CommonLib.WindowControl(RegexName=WindowName).Exists(0, 0):
                        flag = False
                        self.Logfile("[Failed]:App {} Manual Launch".format(Name))
                        return flag
                else:
                    time.sleep(3)
                    if CommonLib.WindowControl(RegexName=WindowName).Exists(0, 0):
                        self.Logfile("[Failed]:APP {} Launch Delay".format(Name))
                        flag = False
                    else:
                        self.Logfile("[PASS]:APP {} Launch Delay".format(Name))
                    time.sleep(20)
            else:
                if Launchdelay != 0:
                    time.sleep(20)
                if not CommonLib.WindowControl(RegexName=WindowName).Exists(0, 0):
                    flag = False
                    self.Logfile("[Failed]:App {} AutoLaunch".format(Name))
                    return flag
                else:
                    self.Logfile("[PASS]:APP {} AutoLaunch".format(Name))
                    self.Logfile("[PASS]:APP {} Launch Delay".format(Name))
            if Maximized:
                if CommonLib.WindowControl(RegexName=WindowName).IsMaximize():
                    self.Logfile("[PASS]:App {} Maximized".format(Name))
                else:
                    self.Logfile("[Failed]:App {} Maximized".format(Name))
                    flag = False
            if Persistent:
                CommonLib.WindowControl(RegexName=WindowName).Close()
                time.sleep(3)
                if Launchdelay == "0":
                    if not CommonLib.WindowControl(RegexName=WindowName).Exists(0, 0):
                        flag = False
                        self.Logfile("[Failed]:App {} Persistent".format(Name))
                        return flag
                    else:
                        self.Logfile("[PASS]:App {} Persistent".format(Name))
                        # self.Logfile("[PASS]:App {} Not AutoDelay".format(Name))
                else:
                    time.sleep(5)
                    if CommonLib.WindowControl(RegexName=WindowName).Exists(0, 0):
                        flag = False
                        self.Logfile("[Failed]:APP {} AutoDelay".format(Name))
            else:
                CommonLib.WindowControl(RegexName=WindowName).Close()
                time.sleep(8)
                if CommonLib.WindowControl(RegexName=WindowName).Exists(0, 0):
                    flag = False
                    self.Logfile("[Failed]:App {} No Persistent".format(Name))
                    return flag
                else:
                    self.Logfile("[PASS]:App {} No Persistent".format(Name))
            return flag
        except:
            self.Logfile("[Failed]:App {} App Check\nErrors:\n{}\n".format(Name, traceback.format_exc()))
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
            for app_path in self.appPath:
                if os.path.exists(app_path):
                    EasyshellLib.CommonUtils.LaunchAppFromFile(app_path)
                    break
                else:
                    continue
            for t in range(50):
                if not EasyshellLib.getElement('MAIN_WINDOW').Exists(searchIntervalSeconds=1):
                    print(t)
                    t += 1
                    time.sleep(2)
                    continue
                else:
                    print('app get window')
                    break
            if not EasyshellLib.getElement('KioskMode').Exists(0, 0):
                self.Logfile('EasyShell is not launch correctly')
                return False
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
            for app_path in self.appPath:
                if os.path.exists(app_path):
                    EasyshellLib.CommonUtils.LaunchAppFromFile(app_path)
                    break
                else:
                    continue
            # Wait HP easy shell launch
            time.sleep(3)
            for t in range(10):
                if not EasyshellLib.getElement('MAIN_WINDOW').Exists(1, 1):
                    time.sleep(1)
                    continue
                else:
                    break
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
            return False


class Shell_Websites(EasyShellTest):
    def __init__(self):
        EasyShellTest.__init__(self)
        self.section_name = 'createWebsites'

    #   ---------------- Website Creation --------------------------
    def CreateWebsite(self, profile):
        try:
            test = self.sections['createWebsites'][profile]
            Name = test["Name"]
            Address = test['Address']
            DefaultHome = test['DefaultHome']
            UseIE = test['UseIE']
            IEFullScreen = test['IEFullScreen']
            EmbedIE = test['EmbedIE']
            AllCloseEmbedIE = test['AllCloseEmbedIE']
            try:
                for app_path in self.appPath:
                    if os.path.exists(app_path):
                        EasyshellLib.CommonUtils.LaunchAppFromFile(app_path)
                        break
                    else:
                        continue
                self.Logfile('---------------Begin to Create website------------')
                for t in range(10):
                    if not EasyshellLib.getElement('MAIN_WINDOW').Exists(1, 1):
                        continue
                    else:
                        print('web Get windows')
                        return False
                EasyshellLib.getElement('KioskMode').Enable()
                EasyshellLib.getElement('DisplayTitle').Enable()
                EasyshellLib.getElement('DisplayWebsites').Enable()
                EasyshellLib.getElement('DisplayBrowser').Enable()
                EasyshellLib.getElement('DisplayAddress').Enable()
                EasyshellLib.getElement('DisplayNavigation').Enable()
                EasyshellLib.getElement('DisplayHome').Enable()
                EasyshellLib.getElement('WebSites').Click()
                if self.Utils(profile, 'Exist'):
                    self.Utils(profile, 'Delete')
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
                    self.Utils(profile, 'default')
                EasyshellLib.getElement('APPLY').Click()
                self.Logfile('[PASS] Create website {} test pass'.format(Name))
                return True
            except:
                self.Logfile("[FAIL]:Website {} Create\nErrors:\n{}\n".format(Name, traceback.format_exc()))
                return False
        except:
            self.Logfile("[FAIL]:Website {} Create\nErrors:\n{}\n".format(Name, traceback.format_exc()))
            return False

    # ---------------- Website Check --------------------------------------
    def CheckWebsite(self, profile):
        flag = True
        test = CommonLib.YmlUtils(os.path.join(self.data, "easyshell_testdata.yaml")).get_item()
        test = test['createWebsites'][profile]
        DefaultHome = test['DefaultHome']
        UseIE = test['UseIE']
        IEFullScreen = test['IEFullScreen']
        EmbedIE = test['EmbedIE']
        AllCloseEmbedIE = test['AllCloseEmbedIE']
        EmbaedPaneName = test['EmbaedPaneName']
        if CommonLib.WindowControl(RegexName='.*- Internet Explorer').Exists(0, 0):
            CommonLib.WindowControl(RegexName='.*- Internet Explorer').Close()
        EasyshellLib.getElement('UserTitles').Click()
        if not self.Utils(profile, 'launch'):
            self.Logfile('[Fail] check website {} error:Launch website fail'.format(profile))
            return False
        time.sleep(5)
        if not UseIE:
            if CommonLib.PaneControl(RegexName=EmbaedPaneName).Exists(0, 0) and not EasyshellLib.getElement(
                    'AddressBar').IsOffscreen:
                if DefaultHome:
                    EasyshellLib.getElement('WebHome').Click()
                    time.sleep(5)
                    if CommonLib.PaneControl(RegexName=EmbaedPaneName).Exists(0, 0):
                        self.Logfile("[PASS]: Websites {} Check".format(profile))
                    else:
                        flag = False
                        self.Logfile("[FAIL]: Websites {} Check".format(profile))
                else:
                    self.Logfile("[PASS]: Websites {} Check".format(profile))
            else:
                flag = False
                self.Logfile("[FAIL]: Websites {} Check".format(profile))
        elif UseIE and not IEFullScreen:
            if CommonLib.WindowControl(RegexName=EmbaedPaneName).Exists(0, 0) and \
                    CommonLib.WindowControl(RegexName=EmbaedPaneName).PaneControl(AutomationId='41477').Exists():
                self.Logfile("[PASS]: Websites {} Check".format(profile))
            else:
                flag = False
                self.Logfile("[FAIL]: Websites {} Check".format(profile))
        elif UseIE and IEFullScreen and not EmbedIE:
            if CommonLib.WindowControl(RegexName=EmbaedPaneName).Exists(0, 0) and \
                    not (
                    CommonLib.WindowControl(RegexName=EmbaedPaneName).PaneControl(AutomationId='41477').Exists(0, 0)):
                self.Logfile("[PASS]: Websites {} Check".format(profile))
            else:
                flag = False
                self.Logfile("[FAIL]: Websites {} Check(110)".format(profile))
        elif UseIE and IEFullScreen and EmbedIE and not AllCloseEmbedIE:
            if CommonLib.PaneControl(RegexName=EmbaedPaneName).Exists(0, 0) and \
                    EasyshellLib.getElement('AddressBar').IsOffscreen \
                    and EasyshellLib.getElement('WebIEClose').IsOffscreen:
                if DefaultHome:
                    EasyshellLib.getElement('WebHome').Click()
                    time.sleep(5)
                    if CommonLib.PaneControl(RegexName=EmbaedPaneName).Exists(0, 0):
                        self.Logfile("[PASS]: Websites {} Check".format(profile))
                    else:
                        flag = False
                        self.Logfile("[FAIL]: Websites {} Home Check(1110)".format(profile))
                else:
                    self.Logfile("[PASS]: Websites {} Check".format(profile))
            else:
                flag = False
                self.Logfile("[FAIL]: Websites {} Check(1110)".format(profile))
        elif UseIE and IEFullScreen and EmbedIE and AllCloseEmbedIE:
            if CommonLib.PaneControl(RegexName=EmbaedPaneName).Exists(0, 0) and \
                    EasyshellLib.getElement('AddressBar').IsOffscreen \
                    and not EasyshellLib.getElement('WebIEClose').IsOffscreen:
                if DefaultHome:
                    EasyshellLib.getElement('WebHome').Click()
                    time.sleep(5)
                    if CommonLib.PaneControl(RegexName=EmbaedPaneName).Exists(0, 0):
                        self.Logfile("[PASS]: Websites {} Check".format(profile))
                    else:
                        flag = False
                        self.Logfile("[FAIL]: Websites {} Check(1111)".format(profile))
                else:
                    self.Logfile("[PASS]: Websites {} Check".format(profile))
            else:
                flag = False
                self.Logfile("[FAIL]: Websites {} Check(1111)".format(profile))
        else:
            flag = False
            self.Logfile("[FAIL]: Websites {} Parameter Error!".format(profile))
        return flag

    def EditModifyWebsite(self, newProfile, oldProfile):
        """
        Make sure Home Default is the same with OldProfile
        """
        with open(os.path.join(self.data, "easyshell_testdata.yaml")) as f:
            test = CommonLib.yaml.safe_load(f)
            new = test['createWebsites'][newProfile]
            old = test['createWebsites'][oldProfile]
            newName = new["Name"]
            newAddress = new['Address']
            newUseIE = new['UseIE']
            oldUseIE = old['UseIE']
            newIEFullScreen = new['IEFullScreen']
            oldIEFullScreen = old['IEFullScreen']
            newEmbedIE = new['EmbedIE']
            oldEmbedIE = old['EmbedIE']
            newAllCloseEmbedIE = new['AllCloseEmbedIE']
            oldAllCloseEmbedIE = old['AllCloseEmbedIE']
        try:
            self.Logfile('-----------Begin to Edit website -------------')
            for app_path in self.appPath:
                if os.path.exists(app_path):
                    EasyshellLib.CommonUtils.LaunchAppFromFile(app_path)
                    break
                else:
                    continue
            for t in range(10):
                if not EasyshellLib.getElement('MAIN_WINDOW').MAIN_WINDOW.Exists(0, 0):
                    time.sleep(1)
                    continue
                else:
                    break
            EasyshellLib.getElement('WebSites').Click()
            self.Utils(oldProfile, 'Edit')
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
                    EasyshellLib.getElement('APPLY').Click()
                    self.Logfile('[PASS] Edit website with new profile {}'.format(newName))
                    return True
            else:
                if newUseIE:
                    CommonLib.SendKey(CommonLib.Keys.VK_TAB)
                else:
                    CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
                    CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
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
                    EasyshellLib.getElement('APPLY').Click()
                    self.Logfile('[PASS] Edit website with new profile {}'.format(newName))
                    return True
            else:
                if newIEFullScreen:
                    CommonLib.SendKey(CommonLib.Keys.VK_TAB)
                else:
                    CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
                    CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
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
                    EasyshellLib.getElement('APPLY').Click()
                    self.Logfile('[PASS] Edit website with new profile {}'.format(newName))
                    return True
            else:
                if newEmbedIE:
                    CommonLib.SendKey(CommonLib.Keys.VK_TAB)
                else:
                    CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
                    CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                    EasyshellLib.getElement('APPLY').Click()
                    self.Logfile('[PASS] Edit website with new profile {}'.format(newName))
                    return True
            if newAllCloseEmbedIE != oldAllCloseEmbedIE:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                EasyshellLib.getElement('APPLY').Click()
                self.Logfile('[PASS] Edit website with new profile {}'.format(newName))
                return True
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                EasyshellLib.getElement('APPLY').Click()
                self.Logfile('[PASS] Edit website with new profile {}'.format(newName))
                return True
        except:
            self.Logfile(
                '[Fail]: Edit website with new profile {}, error: \n{}'.format(newName, traceback.format_exc()))
            return False

    def Utils(self, profile='', op='exist'):
        time.sleep(3)
        with open(os.path.join(self.data, "easyshell_testdata.yaml")) as f:
            test = CommonLib.yaml.safe_load(f)
            test = test['createWebsites'][profile]
            name = test["Name"]
            if CommonLib.TextControl(Name=name).Exists(1, 1):
                txt = CommonLib.TextControl(Name=name)
            else:
                print("didn't get element")
                return False
            websiteControl = txt.GetParentControl().GetParentControl()
            launch = websiteControl.ButtonControl(AutomationId='launchButton2')
            default = websiteControl.ButtonControl(AutoamtionId='homeButton')
            edit = websiteControl.ButtonControl(AutomationId='editButton')
            delete = websiteControl.ButtonControl(AutomationId='deleteButton')
            if op.upper() == 'LAUNCH':
                launch.Click()
                return True
            elif op.upper() == 'EDIT':
                edit.Click()
                return True
            elif op.upper() == 'DELETE':
                try:
                    delete.Click()
                    EasyshellLib.getElement('DeleteYes').Click()
                    EasyshellLib.getElement('APPLY').Click()
                    return True
                except:
                    self.Logfile("[FAIL]:App {} Delete\nErrors:\n{}\n".format(name, traceback.format_exc()))
                    return False
            elif op.upper() == 'EXIST':
                return True
            elif op.upper() == 'DEFAULT':
                default.Click()
                return True
            else:
                pass


class Shell_StoreFront(EasyShellTest):
    def __init__(self):
        EasyShellTest.__init__(self)
        self.section_name = 'createStoreFront'

    # ----------------- StoreFront connection creation ----------------------------
    def CreateStoreFront(self, profile):
        test = CommonLib.YmlUtils(os.path.join(self.data, "easyshell_testdata.yaml")).get_item()
        test = test['createStoreFront'][profile]
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
            for app_path in self.appPath:
                if os.path.exists(app_path):
                    EasyshellLib.CommonUtils.LaunchAppFromFile(app_path)
                    break
                else:
                    continue
            for t in range(10):
                if not EasyshellLib.getElement('MAIN_WINDOW').Exists(1, 1):
                    continue
                else:
                    print('store get window')
                    break
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
            ClearContent()
            CommonLib.SendKeys(str(ConnectionTimeout))
            CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=5)
            CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
            EasyshellLib.getElement('APPLY').Click()
            self.Logfile('[PASS]: View Connection {} Create'.format(Name))
            return True
        except Exception as e:
            self.Logfile("[FAIL]:Storfront {} Create\nErrors:\n{}\n".format(Name, e))
            return False


class Shell_View(EasyShellTest):
    def __init__(self):
        EasyShellTest.__init__(self)
        self.section_name = 'createView'

    # ----------------- VMWare view connection Creation ----------------------------
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
            for app_path in self.appPath:
                if os.path.exists(app_path):
                    EasyshellLib.CommonUtils.LaunchAppFromFile(app_path)
                else:
                    continue
            EasyshellLib.getElement('MAIN_WINDOW').waitExists(10)
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
            return False


class Shell_RDP(EasyShellTest):
    def __init__(self):
        EasyShellTest.__init__(self)
        self.section_name = 'createRDP'

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
            for app_path in self.appPath:
                if os.path.exists(app_path):
                    EasyshellLib.CommonUtils.LaunchAppFromFile(app_path)
                else:
                    continue
            for t in range(50):
                if not EasyshellLib.getElement('MAIN_WINDOW').Exists(searchIntervalSeconds=1):
                    time.sleep(2)
                    continue
                else:
                    print('app get window')
                    break
            EasyshellLib.getElement('MAIN_WINDOW').Exists(10, 3)
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
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            if persistent == 'OFF' or not persistent:
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                CommonLib.SendKey(CommonLib.Keys.VK_TAB)
            else:
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
                    CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
                    CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
                    return False
                CommonLib.SendKey(CommonLib.Keys.VK_TAB, count=2)
                CommonLib.SendKey(CommonLib.Keys.VK_SPACE)
            EasyshellLib.getElement('APPLY').Click()
            self.Logfile('[PASS]: Create RDP Connection {}'.format(name))
        except:
            self.Logfile('[Fail]: Create RDP Connection {}\n{}'.format(name, traceback.format_exc()))


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
            for app_path in self.appPath:
                if os.path.exists(app_path):
                    EasyshellLib.CommonUtils.LaunchAppFromFile(app_path)
                else:
                    continue
            for t in range(50):
                if not EasyshellLib.getElement('MAIN_WINDOW').Exists(searchIntervalSeconds=1):
                    time.sleep(2)
                    continue
                else:
                    print('app get window')
                    break
            EasyshellLib.getElement('MAIN_WINDOW').Exists(10, 3)
            EasyshellLib.getElement('KioskMode').Enable()
            EasyshellLib.getElement('DisplayTitle').Enable()
            EasyshellLib.getElement('DisplayConnections').Enable()
            EasyshellLib.getElement('Connections').Click()
            if self.utils(profile, 'exist', 'connection'):
                self.utils(profile, 'Delete', 'connection')
            EasyshellLib.getElement('CitrixICAAdd').Click()
            time.sleep(5)
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
            old = self.sections['createCitrix'][oldprofile]
            oldname = old["Name"]
            oldhostname = old['Hostname']
            oldusername = old['Username']
            oldautolaunch = old['Autolaunch']
            oldpersistent = old['Persistent']
            for app_path in self.appPath:
                if os.path.exists(app_path):
                    EasyshellLib.CommonUtils.LaunchAppFromFile(app_path)
                    break
                else:
                    continue
            EasyshellLib.getElement('MAIN_WINDOW').Exists(10, 2)
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
        except:
            self.Logfile('[Failed]: Citrix Connection {} Edit\n{}'.format(oldname, traceback.format_exc()))

    @pysnooper.snoop(EasyShellTest().debug)
    def check(self, profile):
        if EasyshellLib.getElement('MAIN_WINDOW').Exists():
            EasyshellLib.getElement('MAIN_WINDOW').SetFocus()
        if self.utils(profile, 'exist', 'connection'):
            self.Logfile('[PASS]:CitrixICA connection {} Check Exist'.format(profile))
            return True
        else:
            self.Logfile('[FAIL]:CitrixICA connection {} Check Not Exist'.format(profile))
            return False


class TaskSwitcher(EasyShellTest):
    def __init__(self):
        EasyShellTest.__init__(self)

    def prepare(self):
        for app_path in self.appPath:
            if os.path.exists(app_path):
                EasyshellLib.CommonUtils.LaunchAppFromFile(self.appPath)
                break
            else:
                continue
        EasyshellLib.getElement('MAIN_WINDOW').waitExists(10)
        EasyshellLib.getElement('KioskMode').Enable()
        EasyshellLib.getElement('EnableTaskSwitcher').Enable()
        EasyshellLib.getElement('Permanent').Enable()

    def enableSoundIconReadOnly(self):
        EasyshellLib.getElement('DisplaySound').Enable()
        EasyshellLib.getElement('DisplaySoundIconInteraction').Disable()
        EasyshellLib.getElement('APPLY').Click()
        self.Logfile("[PASS]: sound icon read only settings")

    def checkSoundReadOnly(self):
        EasyshellLib.getElement('SoundIcon').Click()
        if EasyshellLib.getElement('SoundAdjust').IsOffscreen:
            self.Logfile("[PASS]: sound value is not shown")
            return True
        else:
            self.Logfile("[Fail]: sound value is shown")
            return False

    def enableSoundInteraction(self):
        EasyshellLib.getElement('DisplaySound').Enable()
        EasyshellLib.getElement('DisplaySoundIconInteraction').Enable()
        EasyshellLib.getElement('APPLY').Click()
        self.Logfile("[PASS]: enable sound interaction")

    def checkSoundInteraction(self):
        """
        判断sound interaction 是否工作
        1. 判断当前音量大小，大于80时降低音量操作，小于80时增加音量操作
        2. 通过鼠标键盘两种操作来测试功能
        """
        EasyshellLib.getElement('SoundIcon').Click()
        if not EasyshellLib.getElement('SoundAdjust').IsOffscreen:
            EasyshellLib.getElement('SoundIcon').Click()
            currentVol = EasyshellLib.getElement('SoundAdjust').AccessibleCurrentValue()
            if int(currentVol) < 80:
                EasyshellLib.getElement('SoundAdjustBar').Drag(10, 0)
                tempVol = EasyshellLib.getElement('SoundAdjust').AccessibleCurrentValue()
                if currentVol != tempVol:
                    self.Logfile("[PASS]: Sound Adjust by Mouse")
                else:
                    self.Logfile("[FAIL]: Sound Adjust by Mouse")
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


if __name__ == '__main__':
    for t in range(50):
        print(t)
        if not EasyshellLib.getElement('MAIN_WINDOW').Exists(searchIntervalSeconds=1):
            time.sleep(2)
            t+=1
            continue
        else:
            print('app get window')
            break
    pass

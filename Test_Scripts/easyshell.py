from Test_Scripts.EasyShell_Lib import *
import traceback


def ClearContent(length=50):
    for i in range(length):
        QAUtils.SendKey(QAUtils.Keys().VK_DELETE, 0.01)


class EasyShellTest:
    """
    profile is a test configuration name, it is a dict for app options, load from easyshell_testdata.yaml
    """

    def __init__(self):
        # ------- test root folder - ------------------
        self.path = file_path
        #  ---------------------------------
        self.log_path = os.path.join(self.path, 'Test_Report')
        self.misc = os.path.join(self.path, 'Misc')
        self.data = os.path.join(self.path, 'Test_Data')
        self.casepath = os.path.join(self.path, 'Test_Suite')
        self.testset = os.path.join(self.path, 'testset.xlsx')
        self.sections = YmlUtils(os.path.join(self.data, "easyshell_testdata.yaml")).get_item()
        self.appPath = self.sections['appPath']['easyShellPath']

    def create(self):
        pass

    def edit(self):
        pass

    def check(self):
        pass

    # ------------------------------ Utils -------------------------------------------
    def Logfile(self, rs):
        TxtUtils(os.path.join(self.log_path, "easyShellLog.txt"),
                 'a').set_msg("[{}]:{}\n".format(time.ctime(), rs))

    def utils(self, profile='', op='exist', item='normal'):
        """
        :param profile:  test profile, [test1,test2,,standardApp...]
        :param op: test option [exist | notexist | edit | delete |launch]
        :param item: specific for connection, if item=connection, connection button element=textcontrol.getparent,
                else element=textcontrol.getparent.getparent
        :return: Bool
        """
        time.sleep(3)
        test = self.sections['createApp'][profile]
        name = test["Name"]
        if op.upper() == 'NOTEXIST':
            if TextMethod(Name=name).Exists(0, 0):
                self.Logfile('Check {}-{} Not Exist Fail'.format(profile, name))
                return False
            else:
                self.Logfile('Check {}-{} Not Exist PASS'.format(profile, name))
                return True
        if TextMethod(Name=name).Exists(0, 0):
            txt = TextMethod(Name=name)
        else:
            print("{}-{} Not Exist".format(profile, name))
            return False
        if 'CONN' in item.upper():
            appControl = txt.GetParentControl()
        else:
            appControl = txt.GetParentControl().GetParentControl()
        launch = txt # type=txtcontrol
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
                getElement('DeleteYes').Click()
                getElement('APPLY').Click()
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
                CommonUtils.LaunchAppFromFile(app_path)
                break
            else:
                continue
        time.sleep(5)
        self.Logfile('---------Begin To Test Modify settings----------')
        getElement('KioskMode').Enable()
        for i in test:
            name = i.split(":")[0].strip()  # setting name
            status = i.split(":")[1].strip()  # setting status on/off
            if status == 'ON':
                try:
                    UserInterface_Dict[name].Enable()
                except:
                    flag = False
                    self.Logfile('[Fail]Button {} Enable\n{}'.format(name, traceback.format_exc()))
            elif status == 'OFF':
                try:
                    UserInterface_Dict[name].Disable()
                except:
                    flag = False
                    self.Logfile('[Fail]Button {} Disable\n{}'.format(name, traceback.format_exc()))
            else:
                flag = False
                self.Logfile('[Fail]Button {} Status in test data is not Correct!'.format(name))
        getElement('APPLY').Click()
        getElement('Exit').Click()
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
        for i in test:
            """
            1. 没有测试总开关关闭但是子开关开启时的状态
            2. Wifi 只测试链接窗口是否弹出，没有测试wifi连接功能
            """
            name = i.split(":")[0].strip()
            status = i.split(":")[1].strip()
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
                    if UserKiosk_Dict['UserTitles'].IsShown():
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'DisplayBrowser':
                    if UserKiosk_Dict['UserBrowser'].IsShown():
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'DisplayAdmin':
                    if UserKiosk_Dict['UserAdmin'].IsShown():
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'DisplayPower':
                    if UserKiosk_Dict['UserPower'].IsShown():
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                # ////////// item for Titles //////////////////////////////////////
                if name == 'DisplayApp':
                    UserKiosk_Dict['UserTitles'].Click()
                    time.sleep(1)
                    if UserKiosk_Dict['UserApp'].IsShown():
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'DisplayConnections':
                    UserKiosk_Dict['UserTitles'].Click()
                    time.sleep(1)
                    if UserKiosk_Dict['UserConnection'].IsShown():
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'DisplayStoreFront':
                    UserKiosk_Dict['UserTitles'].Click()
                    time.sleep(1)
                    if UserKiosk_Dict['UserStoreFront'].IsShown():
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'DisplayWebsites':
                    UserKiosk_Dict['UserTitles'].Click()
                    time.sleep(1)
                    if UserKiosk_Dict['UserWebsites'].IsShown():
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                # ////////item for Web browser ///////////////////////////////
                if name == 'DisplayAddress':
                    UserKiosk_Dict['UserBrowser'].Click()
                    time.sleep(1)
                    if UserKiosk_Dict['AddressBar'].IsShown():
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'DisplayHome':
                    UserKiosk_Dict['UserBrowser'].Click()
                    time.sleep(1)
                    if UserKiosk_Dict['WebHome'].IsShown():
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                # ------------------------Item for Admin power----------------------------------------
                if name == 'AllowLock':
                    UserKiosk_Dict['UserPower'].Click()
                    time.sleep(1)
                    if UserKiosk_Dict['Lock'].IsShown():
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'AllowLogoff':
                    UserKiosk_Dict['UserPower'].Click()
                    time.sleep(1)
                    if UserKiosk_Dict['Logoff'].IsShown():
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'AllowRestart':
                    UserKiosk_Dict['UserPower'].Click()
                    time.sleep(1)
                    if UserKiosk_Dict['Restart'].IsShown():
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'AllowShutDown':
                    UserKiosk_Dict['UserPower'].Click()
                    time.sleep(1)
                    if UserKiosk_Dict['Shutdown'].IsShown():
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                # ----------------- Virtual keyboard --------
                if name == 'DisplayVKeyboard':
                    if UserKiosk_Dict['UserKeyBoard'].IsShown():
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                # ---Label that display mac/time/version... at the bottom of UI -----
                if name == 'DisplayTime':
                    if UserKiosk_Dict['Time'].IsShown():
                        self.Logfile("[PASS]: {} is shown".format(name))
                        real_time = QAUtils.getLocalTime('%H:%M')
                        show_time = UserKiosk_Dict['Time'].Name
                        if real_time.split(':')[0] in show_time:
                            self.Logfile("-->[PASS]: {} real time format".format(name))
                        else:
                            self.Logfile("-->[Fail]: {} real time format".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'DisplayIP':
                    if UserKiosk_Dict['IPAddr'].IsShown():
                        self.Logfile("[PASS]: {} is shown".format(name))
                        print(QAUtils.getNetInfo(), '--------net info')
                        real_ip = QAUtils.getNetInfo()['IP']
                        show_ip = UserKiosk_Dict['IPAddr']
                        if real_ip == show_ip:
                            self.Logfile("-->[PASS]: {} real IP".format(name))
                        else:
                            self.Logfile("-->[Fail]: {} real IP".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'DisplayMAC':
                    if UserKiosk_Dict['MACAddr'].IsShown():
                        self.Logfile("[PASS]: {} is shown".format(name))
                        real_mac = QAUtils.getNetInfo()['MAC']
                        show_mac = UserKiosk_Dict['MACAddr']
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
                    if not UserKiosk_Dict['UserTitles'].IsShown():
                        swTitles = False
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'DisplayBrowser':
                    if not UserKiosk_Dict['UserBrowser'].IsShown():
                        swBrowser = False
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'DisplayAdmin':
                    if not UserKiosk_Dict['UserAdmin'].IsShown():
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'DisplayPower':
                    if not UserKiosk_Dict['UserPower'].IsShown():
                        swPower = False
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                # ////////// item for Titles //////////////////////////////////////
                if name == 'DisplayApp':
                    UserKiosk_Dict['UserTitles'].Click()
                    time.sleep(1)
                    if not UserKiosk_Dict['UserApp'].IsShown():
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'DisplayConnections':
                    UserKiosk_Dict['UserTitles'].Click()
                    time.sleep(1)
                    if not UserKiosk_Dict['UserConnection'].IsShown():
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'DisplayStoreFront':
                    UserKiosk_Dict['UserTitles'].Click()
                    time.sleep(1)
                    if not UserKiosk_Dict['UserStoreFront'].IsShown():
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'DisplayWebsites':
                    UserKiosk_Dict['UserTitles'].Click()
                    time.sleep(1)
                    if not UserKiosk_Dict['UserBrowser'].IsShown():
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                # ////////item for Web browser ///////////////////////////////
                if name == 'DisplayAddress':
                    UserKiosk_Dict['UserBrowser'].Click()
                    time.sleep(1)
                    if not UserKiosk_Dict['AddressBar'].IsShown():
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'DisplayHome':
                    UserKiosk_Dict['UserBrowser'].Click()
                    time.sleep(1)
                    if not UserKiosk_Dict['WebHome'].IsShown():
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                # ------------------------Item for Admin power----------------------------------------
                if name == 'AllowLock':
                    UserKiosk_Dict['UserPower'].Click()
                    time.sleep(1)
                    if not UserKiosk_Dict['Lock'].IsShown():
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'AllowLogoff':
                    UserKiosk_Dict['UserPower'].Click()
                    time.sleep(1)
                    if not UserKiosk_Dict['Logoff'].IsShown():
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'AllowRestart':
                    UserKiosk_Dict['UserPower'].Click()
                    time.sleep(1)
                    if not UserKiosk_Dict['Restart'].IsShown():
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'AllowShutDown':
                    UserKiosk_Dict['UserPower'].Click()
                    time.sleep(1)
                    if not UserKiosk_Dict['Shutdown'].IsShown():
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                # ----------------- Virtual keyboard --------
                if name == 'DisplayVKeyboard':
                    if not UserKiosk_Dict['UserKeyBoard'].IsShown():
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                # ---Label that display mac/time/version... at the bottom of UI -----
                if name == 'DisplayTime':
                    if not UserKiosk_Dict['Time'].IsShown():
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'DisplayIP':
                    if not UserKiosk_Dict['IPAddr'].IsShown():
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'DisplayMAC':
                    if not UserKiosk_Dict['MACAddr'].IsShown():
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
                CommonUtils.LaunchAppFromFile(app_path)
                break
            else:
                continue
        time.sleep(5)
        self.Logfile('---------Begin To Test Modify User settings----------')
        getElement('KioskMode').Enable()
        for i in test:
            name = i.split(":")[0].strip()  # setting name
            status = i.split(":")[1].strip()  # setting status on/off
            if status == 'ON':
                try:
                    UserSettings_Dict[name].Enable()
                except:
                    flag = False
                    self.Logfile('[Fail]Button {} Enable\n{}'.format(name, traceback.format_exc()))
            elif status == 'OFF':
                try:
                    UserSettings_Dict[name].Disable()
                except:
                    flag = False
                    self.Logfile('[Fail]Button {} Disable\n{}'.format(name, traceback.format_exc()))
            else:
                flag = False
                self.Logfile('[Fail]Button {} Status in test data is not Correct!'.format(name))
        getElement('APPLY').Click()
        getElement('Exit').Click()
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
        for i in test:
            """
            1. 没有测试总开关关闭但是子开关开启时的状态
            2. Wifi 只测试链接窗口是否弹出，没有测试wifi连接功能
            """
            name = i.split(":")[0].strip()
            status = i.split(":")[1].strip()
            if status == "ON":
                # ///////判断主按钮是否关闭，如果OFF,子按钮不再检查
                if not swSettings:
                    if name in ['AllowMouse', 'AllowKeyboard', 'AllowDisplay', 'AllowSound', 'AllowRegion',
                                'AllowNetworkConn', 'AllowDateTime', 'AllowEasyAccess', 'AllowIEProperty',
                                'AllowWifiConfig']:
                        continue
                # ------------User Settings -----------------------------------
                if name == 'AllowMouse':
                    UserKiosk_Dict['UserSettings'].Click()
                    time.sleep(1)
                    if UserKiosk_Dict['SysMouseIcon'].IsShown():
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'AllowKeyboard':
                    UserKiosk_Dict['UserSettings'].Click()
                    time.sleep(1)
                    if UserKiosk_Dict['SysKeyboardIcon'].IsShown():
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'AllowDisplay':
                    UserKiosk_Dict['UserSettings'].Click()
                    time.sleep(1)
                    if UserKiosk_Dict['SysDisplayIcon'].IsShown():
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'AllowSound':
                    UserKiosk_Dict['UserSettings'].Click()
                    time.sleep(1)
                    if UserKiosk_Dict['SysSoundIcon'].IsShown():
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'AllowRegion':
                    UserKiosk_Dict['UserSettings'].Click()
                    time.sleep(1)
                    if UserKiosk_Dict['SysRegionIcon'].IsShown():
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'AllowNetworkConn':
                    UserKiosk_Dict['UserSettings'].Click()
                    time.sleep(1)
                    if UserKiosk_Dict['SysNetworkConnIcon'].IsShown():
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'AllowDateTime':
                    UserKiosk_Dict['UserSettings'].Click()
                    time.sleep(1)
                    if UserKiosk_Dict['SysDateTimeIcon'].IsShown():
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'AllowEasyAccess':
                    UserKiosk_Dict['UserSettings'].Click()
                    time.sleep(1)
                    if UserKiosk_Dict['SysEaseAccessCenterIcon'].IsShown():
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'AllowIEProperty':
                    UserKiosk_Dict['UserSettings'].Click()
                    time.sleep(1)
                    if UserKiosk_Dict['SysIEIcon'].IsShown():
                        self.Logfile("[PASS]: {} is shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is not shown".format(name))
                if name == 'AllowWifiConfig':
                    UserKiosk_Dict['UserSettings'].Click()
                    time.sleep(1)
                    if UserKiosk_Dict['SysWirelessIcon'].IsShown():
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
                    UserKiosk_Dict['UserSettings'].Click()
                    time.sleep(1)
                    if not UserKiosk_Dict['SysMouseIcon'].Exists(0, 0):
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'AllowKeyboard':
                    UserKiosk_Dict['UserSettings'].Click()
                    time.sleep(1)
                    if not UserKiosk_Dict['SysKeyboardIcon'].Exists(0, 0):
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'AllowDisplay':
                    UserKiosk_Dict['UserSettings'].Click()
                    time.sleep(1)
                    if not UserKiosk_Dict['SysDisplayIcon'].Exists(0, 0):
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'AllowSound':
                    UserKiosk_Dict['UserSettings'].Click()
                    time.sleep(1)
                    if not UserKiosk_Dict['SysSoundIcon'].Exists(0, 0):
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'AllowRegion':
                    UserKiosk_Dict['UserSettings'].Click()
                    time.sleep(1)
                    if not UserKiosk_Dict['SysRegionIcon'].Exists(0, 0):
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'AllowNetworkConn':
                    UserKiosk_Dict['UserSettings'].Click()
                    time.sleep(1)
                    if not UserKiosk_Dict['SysNetworkConnIcon'].Exists(0, 0):
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'AllowDateTime':
                    UserKiosk_Dict['UserSettings'].Click()
                    time.sleep(1)
                    if not UserKiosk_Dict['SysDateTimeIcon'].Exists(0, 0):
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'AllowEasyAccess':
                    UserKiosk_Dict['UserSettings'].Click()
                    time.sleep(1)
                    if not UserKiosk_Dict['SysEaseAccessCenterIcon'].Exists(0, 0):
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'AllowIEProperty':
                    UserKiosk_Dict['UserSettings'].Click()
                    time.sleep(1)
                    if not UserKiosk_Dict['SysIEIcon'].Exists(0, 0):
                        self.Logfile("[PASS]: {} is not shown".format(name))
                    else:
                        flag = False
                        self.Logfile("[Fail]: {} is shown".format(name))
                if name == 'AllowWifiConfig':
                    UserKiosk_Dict['UserSettings'].Click()
                    time.sleep(1)
                    if not UserKiosk_Dict['SysWirelessIcon'].Exists(0, 0):
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
        UserSettings().ModifySettings('test1')
        for app_path in self.appPath:
            if os.path.exists(app_path):
                print(app_path)
                CommonUtils.LaunchAppFromFile(self.appPath)
                break
            else:
                continue
        time.sleep(5)
        try:
            print(bg)
            if bg == 'Custom':
                file = 'customBG'
                getElement('AllowUserSetting').Enable()
                getElement('EnableCustom').Enable()
                getElement('BGFileLocationButton').GetInvokePattern().Invoke()
                getElement('BGFileURLEdit').SetValue(os.path.join(self.misc, "%s.jpg" % file))
                getElement('BGFileOpen').Click()
                getElement('APPLY').Click()
                self.Logfile('[PASS]: Set {} Background file\n'.format(bg))
                return True
            else:
                getElement('AllowUserSetting').Enable()
                for i in UserSettings_Dict.values():
                    i.Enable()
                getElement('EnableCustom').Disable()
                getElement('SelectTheme').GetInvokePattern().Invoke()
                # -------------------select listitem by match the name----------------
                bgcomb = getElement('BGTheme')
                bgcomb.Click()
                time.sleep(3)
                txt = TextMethod(RegexName='.*%s.*' % bg)
                txt.Click()
                getElement('OK').GetInvokePattern().Invoke()
                getElement('APPLY').Click()
                self.Logfile('[PASS]: Set {} Background file\n'.format(bg))
                return True
        except:
            self.Logfile('[FAIL]: Set {} Background file\n{}\n'.format(bg, traceback.format_exc()))
            return False

    def check_background(self, bg='custom'):
        self.Logfile('---------Begin To Check Background ---------------')
        getElement('UserSettings').Click()
        PaneMethod(AutomationId='mainFrame').CaptureToImage('%s.jpg' % bg)
        rgb1 = CommonUtils.getPicRGB('%s.jpg' % bg)
        rgb2 = CommonUtils.getPicRGB(os.path.join(self.misc, '%s.jpg' % bg))
        compare = CommonUtils.compareByRGB(rgb1, rgb2)
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

    # --------------- Applications Creation --------------------
    def check(self, profile):
        self.Logfile('-------------Begin to Check Application --------------')
        flag = True
        content = YmlUtils(os.path.join(self.data, "easyshell_testdata.yaml")).get_item()
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
            if AdminOnly:
                if self.Utils(profile, 'exist'):
                    flag = False
                    self.Logfile("[Fail]:App {} AdminOnly: {}".format(Name, AdminOnly))
                    return flag
                else:
                    flag = True
                    self.Logfile("[PASS]:App {} AdminOnly:{}".format(Name, AdminOnly))
                    return flag
            else:
                if not self.Utils(profile, 'exist'):
                    flag = False
                    self.Logfile("[Fail]:App {} AdminOnly:{}".format(Name, AdminOnly))
                    return flag
            if HideMissApp:
                if self.Utils(profile, 'exist'):
                    flag = False
                    self.Logfile("[Failed]:App {} Hide Missing App".format(Name))
                    return flag
                else:
                    self.Logfile("[PASS]:App {} Hide Missing App".format(Name))
                    return flag
            if Autolaunch == 0:
                self.Utils(profile, "launch")
                if Launchdelay == "None" or Launchdelay is None:
                    for i in range(5):
                        if WindowMethod(RegexName=WindowName).Exists(0, 0):
                            break
                        else:
                            continue
                    if not WindowMethod(RegexName=WindowName).Exists(0, 0):
                        flag = False
                        self.Logfile("[Failed]:App {} Manual Launch".format(Name))
                        return flag
                else:
                    time.sleep(3)
                    if WindowMethod(RegexName=WindowName).Exists(0, 0):
                        self.Logfile("[Failed]:APP {} Launch Delay".format(Name))
                        flag = False
                    else:
                        self.Logfile("[PASS]:APP {} Launch Delay".format(Name))
                    time.sleep(20)
            else:
                if Launchdelay != "None" or Launchdelay is not None:
                    time.sleep(20)
                if not WindowMethod(RegexName=WindowName).Exists(0, 0):
                    flag = False
                    self.Logfile("[Failed]:App {} AutoLaunch".format(Name))
                    return flag
                else:
                    self.Logfile("[PASS]:APP {} AutoLaunch".format(Name))
                    self.Logfile("[PASS]:APP {} Launch Delay".format(Name))
            if Maximized:
                if WindowMethod(RegexName=WindowName).IsMaximize():
                    self.Logfile("[PASS]:App {} Maximized".format(Name))
                else:
                    self.Logfile("[Failed]:App {} Maximized".format(Name))
                    flag = False
            if Persistent:
                WindowMethod(RegexName=WindowName).Close()
                time.sleep(3)
                if Launchdelay == "None" or Launchdelay is None:
                    if not WindowMethod(RegexName=WindowName).Exists(0, 0):
                        flag = False
                        self.Logfile("[Failed]:App {} Persistent".format(Name))
                        return flag
                    else:
                        self.Logfile("[PASS]:App {} Persistent".format(Name))
                        # self.Logfile("[PASS]:App {} Not AutoDelay".format(Name))
                else:
                    time.sleep(5)
                    if WindowMethod(RegexName=WindowName).Exists(0, 0):
                        flag = False
                        self.Logfile("[Failed]:APP {} AutoDelay".format(Name))
            else:
                WindowMethod(RegexName=WindowName).Close()
                time.sleep(8)
                if WindowMethod(RegexName=WindowName).Exists(0, 0):
                    flag = False
                    self.Logfile("[Failed]:App {} No Persistent".format(Name))
                    return flag
                else:
                    self.Logfile("[PASS]:App {} No Persistent".format(Name))
            return flag
        except:
            self.Logfile("[Failed]:App {} App Check\nErrors:\n{}\n".format(Name, traceback.format_exc()))
            return False

    def create(self, profile):
        """
        profile is a test configuration name, it is a dict for app options, load from easyshell_testdata.yaml
        """
        test = self.sections['createApp'][profile]
        Name = test["Name"]
        Paths = test['Path']
        for i in Paths:
            if os.path.exists(i):
                Path = i
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
                    CommonUtils.LaunchAppFromFile(app_path)
                    break
                else:
                    continue
            for i in range(10):
                if not EasyShell_Wnd['MAIN_WINDOW'].Exists(1, 1):
                    time.sleep(1)
                    continue
                else:
                    print('app get window')
                    break
            if not getElement('KiosMode').Exists(0, 0):
                self.Logfile('EasyShell is not launch correctly')
                return False
            getElement('KioskMode').Enable()
            getElement('DisplayTitle').Enable()
            getElement('DisplayApp').Enable()
            getElement('Applications').Click()
            if self.Utils(profile, 'Exist'):
                self.Utils(profile, 'Delete')
            getElement('ApplicationAdd').Click()
            QAUtils.Wait(5)
            QAUtils.SendKeys(Name)
            QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            QAUtils.SendKeys(Path)
            QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            if argument == "None":
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            else:
                QAUtils.SendKeys(argument)
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            if launchdelay == "None":
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            else:
                QAUtils.SendKey(QAUtils.Keys().VK_DELETE)
                QAUtils.SendKeys(str(launchdelay))
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            if autolaunch:
                QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            else:
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            if persistent == 0:
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            else:
                QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            if maximized == 0:
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            else:
                QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            if adminonly == 0:
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            else:
                QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            if hidemissapp == 0:
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            else:
                QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
            getElement('APPLY').Click()
            self.Logfile("[PASS]:App {} Create".format(Name))
            return True
        except:
            self.Logfile("[FAIL]:App {} Create\nErrors:\n{}\n".format(Name, traceback.format_exc()))
            return False

    def edit(self, newProfile, oldProfile):
        try:
            test = YmlUtils(os.path.join(self.data, "easyshell_testdata.yaml")).get_item()
            old = test['createApp'][oldProfile]
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
            new = test['createApp'][newProfile]
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
                    CommonUtils.LaunchAppFromFile(app_path)
                    break
                else:
                    continue
            # Wait HP easy shell launch
            time.sleep(3)
            for i in range(10):
                if not EasyShell_Wnd['MAIN_WINDOW'].Exists(1, 1):
                    time.sleep(1)
                    continue
                else:
                    break
            getElement('Applications').Click()
            # Modify setting//////////////////////////////////////
            self.Utils(oldProfile, 'edit')
            time.sleep(3)
            ClearContent()
            QAUtils.SendKeys(newName)
            QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            ClearContent(len(oldPath) + 5)
            QAUtils.SendKeys(newPath)
            QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            ClearContent()
            if newArgument == 'None':
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            else:
                QAUtils.SendKeys(newArgument)
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            ClearContent()
            if newLaunchdelay == "None":
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            else:
                QAUtils.SendKey(QAUtils.Keys().VK_DELETE)
                QAUtils.SendKeys(str(newLaunchdelay))
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            if oldAutolaunch == newAutolaunch:
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            else:
                QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            if oldPersistent == newPersistent:
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            else:
                QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            if oldMaximized == newMaximized:
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            else:
                QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            if oldAdminonly == newAdminonly:
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            else:
                QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            if oldHidemissapp == newHidemissapp:
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            else:
                QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            QAUtils.SendKey(QAUtils.Keys().VK_TAB, count=2)
            QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
            getElement('APPLY').Click()
            self.Logfile("[PASS]:App {} Edit\n".format(newName))
            return True
        except:
            self.Logfile("[FAIL]:App {} Edit\nErrors:\n{}\n".format(newProfile, traceback.format_exc()))
            return False



class Shell_Websites(EasyShellTest):
    def __init__(self):
        EasyShellTest.__init__(self)

    #   ---------------- Website Creation --------------------------
    def CreateWebsite(self, profile):
        with open(os.path.join(self.data, "easyshell_testdata.yaml")) as f:
            test = yaml.safe_load(f)
            test = test['createWebsites'][profile]
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
                        CommonUtils.LaunchAppFromFile(app_path)
                        break
                    else:
                        continue
                self.Logfile('---------------Begin to Create website------------')
                for i in range(10):
                    if not EasyShell_Wnd['MAIN_WINDOW'].Exists(1, 1):
                        continue
                    else:
                        print('web Get windows')
                        return False
                getElement('KioskMode').Enable()
                UserInterface_Dict['DisplayTitle'].Enable()
                UserInterface_Dict['DisplayWebsites'].Enable()
                UserInterface_Dict['DisplayBrowser'].Enable()
                UserInterface_Dict['DisplayAddress'].Enable()
                UserInterface_Dict['DisplayNavigation'].Enable()
                UserInterface_Dict['DisplayHome'].Enable()
                getElement('WebSites').Click()
                if self.Utils(profile, 'Exist'):
                    self.Utils(profile, 'Delete')
                getElement('WebsiteAdd').Click()
                time.sleep(3)
                print(Name)
                QAUtils.SendKeys(Name)
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
                QAUtils.SendKeys(Address)
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
                if UseIE:
                    QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                    QAUtils.SendKey(QAUtils.Keys().VK_TAB)
                    if IEFullScreen:
                        QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                        QAUtils.SendKey(QAUtils.Keys().VK_TAB)
                        if EmbedIE:
                            QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                            QAUtils.SendKey(QAUtils.Keys().VK_TAB)
                            if AllCloseEmbedIE:
                                QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                                QAUtils.SendKey(QAUtils.Keys().VK_TAB, count=2)
                                QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                            else:
                                # not allcloseembedid |use Id | IE Fullscreen | EmbedIE
                                QAUtils.SendKey(QAUtils.Keys().VK_TAB, count=2)
                                QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                        else:
                            # not embedie | use Ie | Full screen
                            QAUtils.SendKey(QAUtils.Keys().VK_TAB, count=2)
                            QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                    else:
                        # not IE fullscreen | use IE
                        QAUtils.SendKey(QAUtils.Keys().VK_TAB, count=2)
                        QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                else:
                    # Not use IE
                    QAUtils.SendKey(QAUtils.Keys().VK_TAB, count=2)
                    QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                if DefaultHome:
                    self.Utils(profile, 'default')
                getElement('APPLY').Click()
                self.Logfile('[PASS] Create website {} test pass'.format(Name))
                return True
            except:
                self.Logfile("[FAIL]:Website {} Create\nErrors:\n{}\n".format(Name, traceback.format_exc()))
                return False

    # ---------------- Website Check --------------------------------------
    def CheckWebsite(self, profile):
        flag = True
        test = YmlUtils(os.path.join(self.data, "easyshell_testdata.yaml")).get_item()
        test = test['createWebsites'][profile]
        Address = test['Address']
        DefaultHome = test['DefaultHome']
        UseIE = test['UseIE']
        IEFullScreen = test['IEFullScreen']
        EmbedIE = test['EmbedIE']
        AllCloseEmbedIE = test['AllCloseEmbedIE']
        EmbaedPaneName = test['EmbaedPaneName']
        if WindowMethod(RegexName='.*- Internet Explorer').Exists(0, 0):
            WindowMethod(RegexName='.*- Internet Explorer').Close()
        getElement('UserTitles').Click()
        if not self.Utils(profile, 'launch'):
            self.Logfile('[Fail] check website {} error:Launch website fail'.format(profile))
            return False
        time.sleep(5)
        if not UseIE:
            if PaneMethod(RegexName=EmbaedPaneName).Exists(0, 0) and getElement('AddressBar').IsShown():
                if DefaultHome:
                    getElement('WebHome').Click()
                    time.sleep(5)
                    if PaneMethod(RegexName=EmbaedPaneName).Exists(0, 0):
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
            if WindowMethod(RegexName=EmbaedPaneName).Exists(0, 0) and \
                    WindowMethod(RegexName=EmbaedPaneName).PaneControl(AutomationId='41477').Exists():
                self.Logfile("[PASS]: Websites {} Check".format(profile))
            else:
                flag = False
                self.Logfile("[FAIL]: Websites {} Check".format(profile))
        elif UseIE and IEFullScreen and not EmbedIE:
            if WindowMethod(RegexName=EmbaedPaneName).Exists(0, 0) and \
                    not (WindowMethod(RegexName=EmbaedPaneName).PaneControl(AutomationId='41477').Exists(0, 0)):
                self.Logfile("[PASS]: Websites {} Check".format(profile))
            else:
                flag = False
                self.Logfile("[FAIL]: Websites {} Check(110)".format(profile))
        elif UseIE and IEFullScreen and EmbedIE and not AllCloseEmbedIE:
            if PaneMethod(RegexName=EmbaedPaneName).Exists(0, 0) and not (getElement('AddressBar').IsShown()) \
                    and not (getElement('WebIEClose').IsShown()):
                if DefaultHome:
                    getElement('WebHome').Click()
                    time.sleep(5)
                    if PaneMethod(RegexName=EmbaedPaneName).Exists(0, 0):
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
            if PaneMethod(RegexName=EmbaedPaneName).Exists(0, 0) and not (getElement('AddressBar').IsShown()) \
                    and getElement('WebIEClose').IsShown():
                if DefaultHome:
                    getElement('WebHome').Click()
                    time.sleep(5)
                    if PaneMethod(RegexName=EmbaedPaneName).Exists(0, 0):
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
            test = yaml.safe_load(f)
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
                    CommonUtils.LaunchAppFromFile(app_path)
                    break
                else:
                    continue
            for i in range(10):
                if not EasyShell_Wnd['MAIN_WINDOW'].Exists(0, 0):
                    time.sleep(1)
                    continue
                else:
                    break
            getElement('WebSites').Click()
            self.Utils(oldProfile, 'Edit')
            time.sleep(3)
            ClearContent()
            QAUtils.SendKeys(newName)
            QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            ClearContent()
            QAUtils.SendKeys(newAddress)
            QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            if newUseIE != oldUseIE and newUseIE:
                if newUseIE:
                    QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                    QAUtils.SendKey(QAUtils.Keys().VK_TAB)
                else:
                    QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                    QAUtils.SendKey(QAUtils.Keys().VK_TAB, count=2)
                    QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                    getElement('APPLY').Click()
                    self.Logfile('[PASS] Edit website with new profile {}'.format(newName))
                    return True
            else:
                if newUseIE:
                    QAUtils.SendKey(QAUtils.Keys().VK_TAB)
                else:
                    QAUtils.SendKey(QAUtils.Keys().VK_TAB, count=2)
                    QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                    getElement('APPLY').Click()
                    self.Logfile('[PASS] Edit website with new profile {}'.format(newName))
                    return True
            if newIEFullScreen != oldIEFullScreen:
                if newIEFullScreen:
                    QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                    QAUtils.SendKey(QAUtils.Keys().VK_TAB)
                else:
                    QAUtils.SendKey(QAUtils.Keys().VK_TAB)
                    QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                    QAUtils.SendKey(QAUtils.Keys().VK_TAB, count=2)
                    QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                    getElement('APPLY').Click()
                    self.Logfile('[PASS] Edit website with new profile {}'.format(newName))
                    return True
            else:
                if newIEFullScreen:
                    QAUtils.SendKey(QAUtils.Keys().VK_TAB)
                else:
                    QAUtils.SendKey(QAUtils.Keys().VK_TAB, count=2)
                    QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                    getElement('APPLY').Click()
                    self.Logfile('[PASS] Edit website with new profile {}'.format(newName))
                    return True
            if newEmbedIE != oldEmbedIE:
                if newEmbedIE:
                    QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                    QAUtils.SendKey(QAUtils.Keys().VK_TAB)
                else:
                    QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                    QAUtils.SendKey(QAUtils.Keys().VK_TAB, count=2)
                    QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                    getElement('APPLY').Click()
                    self.Logfile('[PASS] Edit website with new profile {}'.format(newName))
                    return True
            else:
                if newEmbedIE:
                    QAUtils.SendKey(QAUtils.Keys().VK_TAB)
                else:
                    QAUtils.SendKey(QAUtils.Keys().VK_TAB, count=2)
                    QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                    getElement('APPLY').Click()
                    self.Logfile('[PASS] Edit website with new profile {}'.format(newName))
                    return True
            if newAllCloseEmbedIE != oldAllCloseEmbedIE:
                QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                QAUtils.SendKey(QAUtils.Keys().VK_TAB, count=2)
                QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                getElement('APPLY').Click()
                self.Logfile('[PASS] Edit website with new profile {}'.format(newName))
                return True
            else:
                QAUtils.SendKey(QAUtils.Keys().VK_TAB, count=2)
                QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                getElement('APPLY').Click()
                self.Logfile('[PASS] Edit website with new profile {}'.format(newName))
                return True
        except:
            self.Logfile('[Fail]: Edit website with new profile {}, error: \n{}'.format(newName, traceback.format_exc()))
            return False

    def Utils(self, profile='', op='exist'):
        time.sleep(3)
        with open(os.path.join(self.data, "easyshell_testdata.yaml")) as f:
            test = yaml.safe_load(f)
            test = test['createWebsites'][profile]
            name = test["Name"]
            if TextMethod(Name=name).Exists(1, 1):
                txt = TextMethod(Name=name)
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
            elif op .upper() == 'EDIT':
                edit.Click()
                return True
            elif op.upper() == 'DELETE':
                try:
                    delete.Click()
                    getElement('DeleteYes').Click()
                    getElement('APPLY').Click()
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

    # ----------------- StoreFront connection creation ----------------------------
    def CreateStoreFront(self, profile):
        test = YmlUtils(os.path.join(self.data, "easyshell_testdata.yaml")).get_item()
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
                    CommonUtils.LaunchAppFromFile(app_path)
                    break
                else:
                    continue
            for i in range(10):
                if not EasyShell_Wnd['MAIN_WINDOW'].Exists(1, 1):
                    continue
                else:
                    print('store get window')
                    break
            getElement('KioskMode').Enable()
            getElement('DisplayTitle').Enable()
            getElement('DisplayStoreFront').Enable()
            getElement('StoreFront').Click()
            if self.Utils(profile, 'Exist'):
                self.Utils(profile, 'Delete')
            getElement('StoreFrontAdd').Click()
            time.sleep(5)
            QAUtils.SendKeys(Name)
            QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            QAUtils.SendKeys(URL)
            QAUtils.SendKey(QAUtils.Keys().VK_TAB, count=2)
            QAUtils.SendKeys(LogonMethod)
            QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            if HideDomain == 'OFF':
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            else:
                QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            QAUtils.SendKeys(Username)
            QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            QAUtils.SendKeys(Password)
            QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            QAUtils.SendKeys(Domain)
            QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            if Autolaunch == 'OFF':
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            else:
                QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            if not SelectStore:
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            else:
                QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                time.sleep(1)
                QAUtils.SendKeys(StoreName)
                QAUtils.SendKey(QAUtils.Keys().VK_TAB, count=2)
                QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                time.sleep(3)
                QAUtils.SendKey(QAUtils.Keys().VK_TAB, count=3)
                QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                QAUtils.SendKey(QAUtils.Keys().VK_TAB, count=2)
                QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            ClearContent()
            QAUtils.SendKeys(str(Launchdelay))
            QAUtils.SendKey(QAUtils.Keys().VK_TAB, count=2)
            if CustomLogon == 'None':
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            else:
                QAUtils.SendKeys(CustomLogon)
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            QAUtils.SendKey(QAUtils.Keys().VK_TAB, count=2)
            QAUtils.SendKey(QAUtils.Keys().VK_RIGHT)
            QAUtils.SendKey(QAUtils.Keys().VK_TAB, count=3)
            if DesktopToolbar == 'OFF':
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            else:
                QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            ClearContent()
            QAUtils.SendKeys(str(ConnectionTimeout))
            QAUtils.SendKey(QAUtils.Keys().VK_TAB, count=5)
            QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
            getElement('APPLY').Click()
            self.Logfile('[PASS]: View Connection {} Create'.format(Name))
            return True
        except Exception as e:
            self.Logfile("[FAIL]:Storfront {} Create\nErrors:\n{}\n".format(Name, e))
            return False


class Shell_View(EasyShellTest):
    def __init__(self):
        EasyShellTest.__init__(self)

    # ----------------- VMWare view connection Creation ----------------------------
    def CreateView(self, profile):
        test = YmlUtils(os.path.join(self.data, "easyshell_testdata.yaml")).get_item()
        test = test['createView'][profile]
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
                    CommonUtils.LaunchAppFromFile(app_path)
                else:
                    continue
            EasyShell_Wnd['MAIN_WINDOW'].waitExists(10)
            getElement('KioskMode').Enable()
            getElement('DisplayTitle').Enable()
            getElement('DisplayConnections').Enable()
            getElement('Connections').Click()
            if self.Utils(profile, 'exist'):
                self.Utils(profile, 'Delete')
            getElement('VMwareAdd').Click()
            time.sleep(5)
            QAUtils.SendKeys(Name)
            QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            QAUtils.SendKeys(Hostname)
            QAUtils.SendKey(QAUtils.Keys().VK_TAB, count=2)
            ClearContent()
            QAUtils.SendKeys(str(Launchdelay))
            QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            if Argument == 'None':
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            else:
                QAUtils.SendKeys(Argument)
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            if Autolaunch == 'OFF' or not Autolaunch:
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            else:
                QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            if Persistent == 'OFF' or not Persistent:
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            else:
                QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            QAUtils.SendKey(QAUtils.Keys().VK_TAB, count=2)
            QAUtils.SendKey(QAUtils.Keys().VK_RIGHT)
            QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            QAUtils.SendKeys(Layout)
            QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            if ConnUSBStartup == 'OFF':
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            else:
                QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            if ConnUSBInsertion == 'OFF':
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            else:
                QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
                QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            QAUtils.SendKeys(Username)
            QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            QAUtils.SendKeys(Password)
            QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            QAUtils.SendKeys(Domain)
            QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            QAUtils.SendKeys(DesktopName)
            QAUtils.SendKey(QAUtils.Keys().VK_TAB)
            QAUtils.SendKey(QAUtils.Keys().VK_SPACE)
            getElement('APPLY').Click()
            self.Logfile('[PASS]: View Connection {} Create'.format(Name))
            return True
        except:
            self.Logfile("[FAIL]: View Connection {} Create\nErrors:\n{}\n".format(Name, traceback.format_exc()))
            return False


class Shell_RDP(EasyShellTest):
    def __init__(self):
        EasyShellTest.__init__(self)


class Shell_Citrix(EasyShellTest):
    def __init__(self):
        EasyShellTest.__init__(self)


class TaskSwitcher(EasyShellTest):
    def __init__(self):
        EasyShellTest.__init__(self)

    def prepare(self):
        for app_path in self.appPath:
            if os.path.exists(app_path):
                CommonUtils.LaunchAppFromFile(self.appPath)
                break
            else:
                continue
        EasyShell_Wnd['MAIN_WINDOW'].waitExists(10)
        getElement('KioskMode').Enable()
        getElement('EnableTaskSwitcher').Enable()
        getElement('Permanent').Enable()

    def enableSoundIconReadOnly(self):
        getElement('DisplaySound').Enable()
        getElement('DisplaySoundIconInteraction').Disable()
        getElement('APPLY').Click()
        self.Logfile("[PASS]: sound icon read only settings")

    def checkSoundReadOnly(self):
        getElement('SoundIcon').Click()
        if not getElement('SoundAdjust').IsShown():
            self.Logfile("[PASS]: sound value is not shown")
            return True
        else:
            self.Logfile("[Fail]: sound value is shown")
            return False

    def enableSoundInteraction(self):
        getElement('DisplaySound').Enable()
        getElement('DisplaySoundIconInteraction').Enable()
        getElement('APPLY').Click()
        self.Logfile("[PASS]: enable sound interaction")

    def checkSoundInteraction(self):
        """
        判断sound interaction 是否工作
        1. 判断当前音量大小，大于80时降低音量操作，小于80时增加音量操作
        2. 通过鼠标键盘两种操作来测试功能
        """
        getElement('SoundIcon').Click()
        if getElement('SoundAdjust').IsShown():
            getElement('SoundIcon').Click()
            currentVol = getElement('SoundAdjust').AccessibleCurrentValue()
            if int(currentVol) < 80:
                getElement('SoundAdjustBar').Drag(10, 0)
                tempVol = getElement('SoundAdjust').AccessibleCurrentValue()
                if currentVol != tempVol:
                    self.Logfile("[PASS]: Sound Adjust by Mouse")
                else:
                    self.Logfile("[FAIL]: Sound Adjust by Mouse")
                    return False
                QAUtils.SendKey(QAUtils.Keys().VK_RIGHT)
                finalVol = getElement('SoundAdjust').AccessibleCurrentValue()
                if finalVol != tempVol:
                    self.Logfile("[PASS]: Sound Adjust by Keyboard")
                    return True
                else:
                    self.Logfile("[FAIL]: Sound Adjust by Keyboard")
                    return False
            else:
                getElement('SoundAdjustBar').Drag(-10, 0)
                tempVol = getElement('SoundAdjust').AccessibleCurrentValue()
                if currentVol != tempVol:
                    self.Logfile("[PASS]: Sound Adjust by Mouse")
                else:
                    self.Logfile("[FAIL]: Sound Adjust by Mouse")
                    return False
                QAUtils.SendKey(QAUtils.Keys().VK_LEFT)
                finalVol = getElement('SoundAdjust').AccessibleCurrentValue()
                if finalVol != tempVol:
                    self.Logfile("[PASS]: {} Sound Adjust by Keyboard")
                    return True
                else:
                    self.Logfile("[FAIL]: {} Sound Adjust by Keyboard")
                    return False
        else:
            self.Logfile("[Fail]: sound value is not shown")
            return False

            # --------------- Task Switcher ----------------------------------
            '''
            if name == 'DisplaySwitcherTime':
                if UserKiosk_Dict['SwitcherTime'].IsShown():
                    self.Logfile("[PASS]: {} is shown".format(name))
                else:
                    flag = False
                    self.Logfile("[Fail]: {} is not shown".format(name))
            if name == 'DisplaySound':
                if UserKiosk_Dict['SoundIcon'].IsShown():
                    self.Logfile("[PASS]: {} is shown".format(name))
                else:
                    flag = False
                    self.Logfile("[Fail]: {} is not shown".format(name))

            if name == 'DisplaySoundIconInteraction':
                """
                判断sound interaction 是否工作
                1. 判断当前音量大小，大于80时降低音量操作，小于80时增加音量操作
                2. 通过鼠标键盘两种操作来测试功能
                """
                if UserKiosk_Dict['SoundIcon'].IsShown():
                    UserKiosk_Dict['SoundIcon'].Click()
                    currentVol = ButtonList.SoundAdjust().AccessibleCurrentValue()
                    if int(currentVol) < 80:
                        ButtonList.SoundAdjustBar().Drag(10, 0)
                        tempVol = ButtonList.SoundAdjust().AccessibleCurrentValue()
                        if currentVol != tempVol:
                            self.Logfile("[PASS]: {} Sound Adjust by Mouse".format(name))
                        else:
                            flag = False
                            self.Logfile("[FAIL]: {} Sound Adjust by Mouse".format(name))
                        QAUtils.SendKey(QAUtils.Keys().VK_RIGHT)
                        finalVol = ButtonList.SoundAdjust().AccessibleCurrentValue()
                        if finalVol != tempVol:
                            self.Logfile("[PASS]: {} Sound Adjust by Keyboard".format(name))
                        else:
                            flag = False
                            self.Logfile("[FAIL]: {} Sound Adjust by Keyboard".format(name))
                    else:
                        ButtonList.SoundAdjustBar().Drag(-10, 0)
                        tempVol = ButtonList.SoundAdjust().AccessibleCurrentValue()
                        if currentVol != tempVol:
                            self.Logfile("[PASS]: {} Sound Adjust by Mouse".format(name))
                        else:
                            flag = False
                            self.Logfile("[FAIL]: {} Sound Adjust by Mouse".format(name))
                        QAUtils.SendKey(QAUtils.Keys().VK_LEFT)
                        finalVol = ButtonList.SoundAdjust().AccessibleCurrentValue()
                        if finalVol != tempVol:
                            self.Logfile("[PASS]: {} Sound Adjust by Keyboard".format(name))
                        else:
                            flag = False
                            self.Logfile("[FAIL]: {} Sound Adjust by Keyboard".format(name))
                else:
                    continue
            '''



if __name__ == '__main__':
    # Shell_Application().CheckApp('test1')
    # Shell_StoreFront().CreateStoreFront('standardStore')
    pass
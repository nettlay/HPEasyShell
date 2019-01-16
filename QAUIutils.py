import socket
from psutil import net_if_addrs
import re
import uiautomation
import os, time
import win32api,win32con
import subprocess

'''
Some element might mistake recognize automationId or Name:
Minimize: Only Name is useful, AutomationId actually is "" even though we can get the value
'''


class QATools:
    @staticmethod
    def getNetInfo():
        ip_addr = ''
        mac_addr = ''
        ip_pattern = r'(([01]{0,1}\d{0,1}\d|2[0-4]\d|25[0-5])\.){3}(?:[01]{0,1}\d{0,1}\d|2[0-4]\d|25[0-5])'
        for k, v in net_if_addrs().items():
            if 'Ethernet' in k:
                for item in v:
                    if '-' in item[1] and len(item[1]) == 17:
                        mac_addr = item[1]
                    if re.match(ip_pattern, item[1]):
                        ip_addr = item[1]
                return {'MAC': mac_addr, 'IP': ip_addr}

    @staticmethod
    def getLocalTime(fmt='%I:%M:%S %Y/%m/%d', fullTime=False):
        if fullTime:
            # fmt = '%H:%M:%S %Y/%m/%d %Z'
            hour = int(time.strftime('%H'))
        else:
            hour = int(time.strftime('%I'))
        minute = int(time.strftime('%M'))
        second = int(time.strftime('%S'))
        if '%Y' in fmt:
            year = int(time.strftime('%Y'))
        elif '%y' in fmt:
            year = int(time.strftime('%y'))
        else:
            year = ''
        month = int(time.strftime('%m'))
        day = int(time.strftime('%d'))
        time1 = fmt
        if '%I' in fmt:
            time1 = time1.replace("%I", str(hour))
        elif '%H' in fmt:
            time1 = time1.replace('%H', str(hour))
        else:
            pass
        if '%M' in fmt:
            time1 = time1.replace('%M', str(minute))
        if '%S' in fmt:
            time1 = time1.replace('%S', str(second))
        if '%Y' in fmt:
            time1 = time1.replace('%Y', str(year))
        elif '%y' in fmt:
            time1 = time1.replace('%y', str(year))
        else:
            pass
        if '%m' in fmt:
            time1 = time1.replace('%m', str(month))
        if '%d' in fmt:
            time1 = time1.replace('%d', str(day))
        return time1

    @staticmethod
    def addUser(username, password, group='users'):
        # Default group is Users
        os.system('net user /delete {}'.format(username))
        os.system('net user /add {} {}'.format(username, password))
        if group == 'Administrators':
            os.system('net localgroup Administrators {} /add'.format(username))
            return True
        else:
            return True

    @staticmethod
    def getSid():
        """
        Get the sessionId for all users in windows
        :return: sessionId key,value dictionary userlist
        """
        userList = {}
        t = os.popen('wmic useraccount get name,sid').read()
        sidInfo = t.split('\n')
        for i in sidInfo:
            if i == "":
                continue
            ls = i.split(' S')
            if ls[0].strip() == 'Name':
                continue
            userList[ls[0].strip()] = 'S{}'.format(ls[1]).strip()
        return userList

    @staticmethod
    def IsShown(element):
        if element.IsOffScreen:
            return False
        else:
            return True

    @staticmethod
    def Wait(waitTime=3):
        time.sleep(waitTime)

    @staticmethod
    def SendKeys(keys, waitTime=0.1):
        return uiautomation.SendKeys(keys, waitTime=waitTime)

    @staticmethod
    def SendKey(key, waitTime=0.1):
        return uiautomation.SendKey(key, waitTime=waitTime)

    @staticmethod
    def Keys():
        return uiautomation.Keys

    @staticmethod
    def ControlFromCursor():
        return uiautomation.ControlFromCursor()

    @staticmethod
    def CheckWindow(name):
        time.sleep(3)
        if Method(RegexName=name).Exists(0, 0):
            return True
        else:
            return False

    @staticmethod
    def GetWindow(name):
        time.sleep(3)
        if Method(RegexName=name).Exists(0, 0):
            return WindowMethod(RegexName=name)
        else:
            return False

    @staticmethod
    def GetListFromFile(file):
        """
        Read file and return list contains all the lines
        :param file: file
        :return: file_list
        """
        file_list=[]
        with open(file) as f:
            rs = f.readlines()
            for line in rs:
                file_list.append(line.strip())
        return file_list

    @staticmethod
    def Reboot(waitTime=3):
        os.system('shutdown -r -t {}'.format(waitTime))
        time.sleep(3)

    @staticmethod
    def SwitchUser(username="Admin",password="Admin",domain=""):
        """
        Before program, Please Add HKLM\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon
        To Registry exclusions of Write Filter
        """
        root = win32con.HKEY_LOCAL_MACHINE
        key = win32api.RegOpenKeyEx(root,'SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon', 0,
                                    win32con.KEY_ALL_ACCESS | win32con.KEY_WOW64_64KEY | win32con.KEY_WRITE)
        win32api.RegSetValueEx(key,"DefaultUserName", 0, win32con.REG_SZ, username)
        win32api.RegSetValueEx(key, "DefaultPassWord", 0, win32con.REG_SZ, password)
        if domain != "":
            win32api.RegSetValueEx(key, "DefaultDomain", 0,  win32con.REG_SZ, domain)

    @staticmethod
    def Write2File(filename, src):
        with open(filename, 'a', encoding='utf8') as f:
            f.write(src)

    @staticmethod
    def LaunchAppFromControl(name):
        """
        Start application via Control panel
        :param name: app name shown in control panel
        :return: None
        """
        os.system("control panel")

        for i in range(10):
            print(i)
            if not WindowMethod(Name="All Control Panel Items").Exists(0,0):
                print("not exist")
                continue
            else:
                mainWnd = WindowMethod(Name="All Control Panel Items")
                break

        mainWnd.Maximize()
        ButtonMethod(AutomationId="ActionButton").Click()
        Method(AutomationId="51").Click()
        controlPane = mainWnd.PaneControl(AutomationId='CategoryPanel')
        items = controlPane.GetChildren()
        for item in items:
            if item.Name == name:
                item.Invoke()
                time.sleep(2)
                break
        if WindowMethod(Name="All Control Panel Items").Exists():
            WindowMethod(Name="All Control Panel Items").Close()

    @staticmethod
    def GetInstallationList(filename):
        """
        Get installation application in program and features, write to filename
        :param filename: filename
        :return: None
        """
        os.system('control panel')
        mainWnd = WindowMethod(Name="All Control Panel Items")
        mainWnd.Maximize()
        controlPane = mainWnd.PaneControl(AutomationId='CategoryPanel')
        items = controlPane.GetChildren()
        for item in items:
            if item.Name == "Programs and Features":
                item.Invoke()
        redirectWnd = WindowMethod(Name="Programs and Features")
        installationPool = redirectWnd.ListControl(AutomationId="1")
        items = installationPool.GetChildren()
        for item in items:
            if item.ControlType == uiautomation.ControlType.ScrollBarControl:
                continue
            subitems = item.GetChildren()
            QATools.Write2File(filename, "{}:{}\n".format(subitems[0].Name, subitems[4].Name))
        redirectWnd.Close()

    @staticmethod
    def LaunchAppFromFile(path):
        subprocess.Popen(path)
        time.sleep(1)

    @staticmethod
    def addLocalRegValue(path, keyName, value, keyType='SZ'):
        root = win32con.HKEY_LOCAL_MACHINE
        # os.getlogin() only for python3, can be replaced with getpass.getuser()
        try:
            key = win32api.RegOpenKeyEx(root, path, 0, win32con.KEY_ALL_ACCESS | win32con.KEY_WOW64_64KEY)
            if keyType.upper() == 'SZ':
                key_type = win32con.REG_SZ
            elif keyType.upper() == 'BIN':
                key_type = win32con.REG_BINARY
            elif keyType.upper() == 'DW':
                key_type = win32con.REG_DWORD
            else:
                key_type = win32con.REG_SZ
            win32api.RegSetValueEx(key, keyName, 0, key_type, value)
        except Exception as e:
            raise ("Key not found :\n{}".format(e))

    @staticmethod
    def addUserRegValue(path, keyName, value, keyType='SZ'):
        root = win32con.HKEY_USERS
        # os.getlogin() only for python3, can be replaced with getpass.getuser()
        path = '{}\\{}'.format(QATools.getSid()[os.getlogin()], path)
        try:
            key = win32api.RegOpenKeyEx(root, path, 0 , win32con.KEY_ALL_ACCESS | win32con.KEY_WOW64_64KEY)
            if keyType.upper() == 'SZ':
                key_type = win32con.REG_SZ
            elif keyType.upper() == 'BIN':
                key_type = win32con.REG_BINARY
            elif keyType.upper() == 'DW':
                key_type = win32con.REG_DWORD
            else:
                key_type = win32con.REG_SZ
            win32api.RegSetValueEx(key, keyName, 0, key_type, value)
        except Exception as e:
            raise("Key not found :\n{}".format(e))


class Method(uiautomation.Control):
    def __init__(self, element=0, searchFromControl=None, searchDepth=0xFFFFFFFF, searchWaitTime=0.5, foundIndex=1,
                 **searchPorpertyDict):
        uiautomation.Control.__init__(self, element, searchFromControl, searchDepth,
                                      searchWaitTime, foundIndex, **searchPorpertyDict)

    def Drag(self, x=0, y=0):
        x1, y1 = self.BoundingRectangle[0], self.BoundingRectangle[1]
        print(x1, y1)
        uiautomation.DragDrop(x1, y1, x1+x, y1+y)

    def IsShown(self):
        if self.IsOffScreen:
            return False
        else:
            return True


class ButtonMethod(uiautomation.ButtonControl, Method):
    def __init__(self, element=0, searchFromControl=None, searchDepth=0xFFFFFFFF, searchWaitTime=0.5, foundIndex=1,
                 **searchPorpertyDict):
        uiautomation.ButtonControl.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex,
                                            **searchPorpertyDict)

    def ClickEx(self, waitTime=0.1):
        """
        Toggle or Invoke button with build-in method instead of Click
        """
        if self.IsTogglePatternAvailable():
            self.Toggle(waitTime)
        if self.IsInvokePatternAvailable():
            self.Invoke(waitTime)

    def GetStatus(self):
        if not self.Exists():
            print("Button not Found, Please Double check your parameter!!")
            return
        if self.IsTogglePatternAvailable():
            return self.CurrentToggleState()
        else:
            return None

    def Enable(self):
        # if not self.Exists(0,0):
        #     print("Button not Found, Please Double check your parameter!!")
        #     return
        if self.IsTogglePatternAvailable():
            if self.CurrentToggleState():
                return
            else:
                self.ClickEx()

    def Disable(self):
        # if not self.Exists():
        #     print("Button not Found, Please Double check your parameter!!")
        #     return
        if self.IsTogglePatternAvailable():
            if self.CurrentToggleState():
                self.ClickEx()
            else:
                return


class CheckBoxMethod(uiautomation.CheckBoxControl, Method):
    def __init__(self, element=0, searchFromControl=None, searchDepth=0xFFFFFFFF, searchWaitTime=0.5, foundIndex=1,
                 **searchPorpertyDict):
        uiautomation.CheckBoxControl.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex,
                                              **searchPorpertyDict)


class DataGridMethod(uiautomation.DataGridControl, Method):
    def __init__(self, element=0, searchFromControl=None, searchDepth=0xFFFFFFFF, searchWaitTime=0.5, foundIndex=1,
                 **searchPorpertyDict):
        uiautomation.DataGridControl.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex,
                                              **searchPorpertyDict)


class DataItemMethod(uiautomation.DataItemControl, Method):
    def __init__(self, element=0, searchFromControl=None, searchDepth=0xFFFFFFFF, searchWaitTime=0.5, foundIndex=1,
                 **searchPorpertyDict):
        uiautomation.DataItemControl.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex,
                                              **searchPorpertyDict)


class PaneMethod(uiautomation.PaneControl, Method):
    def __init__(self, element=0, searchFromControl=None, searchDepth=0xFFFFFFFF, searchWaitTime=0.5, foundIndex=1,
                 **searchPorpertyDict):
        uiautomation.PaneControl.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex,
                                          **searchPorpertyDict)


class TextMethod(uiautomation.TextControl, Method):
    def __init__(self, element=0, searchFromControl=None, searchDepth=0xFFFFFFFF, searchWaitTime=0.5, foundIndex=1,
                 **searchPorpertyDict):
        uiautomation.TextControl.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex,
                                          **searchPorpertyDict)

    def compareWith(self, obj):
        if not self.Exists():
            print("Text Box not Found, Please Double check your parameter!!")
            return
        if self.Name == obj:
            return True
        else:
            return False

    def contains(self, obj):
        if not self.Exists():
            print("Text Box not Found, Please Double check your parameter!!")
            return
        if obj in self.Name:
            return True
        else:
            return False


class EditMethod(uiautomation.EditControl, Method):
    def __init__(self, element=0, searchFromControl=None, searchDepth=0xFFFFFFFF, searchWaitTime=0.5, foundIndex=1,
                 **searchPorpertyDict):
        uiautomation.EditControl.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex,
                                          **searchPorpertyDict)

    def compareWith(self, obj):
        if not self.Exists():
            print("Edit Box not Found, Please Double check your parameter!!")
            return
        if self.CurrentValue() == obj:
            return True
        else:
            return False

    def contains(self, obj):
        if not self.Exists():
            print("Edit Box not Found, Please Double check your parameter!!")
            return
        if obj in self.CurrentValue():
            return True
        else:
            return False


class ListItemMethod(uiautomation.ListItemControl, Method):
    def __init__(self, element=0, searchFromControl=None, searchDepth=0xFFFFFFFF, searchWaitTime=0.5, foundIndex=1,
                 **searchPorpertyDict):
        uiautomation.ListItemControl.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex,
                                              **searchPorpertyDict)


class ListMethod(uiautomation.ListControl, Method):
    def __init__(self, element=0, searchFromControl=None, searchDepth=0xFFFFFFFF, searchWaitTime=0.5, foundIndex=1,
                 **searchPorpertyDict):
        uiautomation.ListControl.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex,
                                          **searchPorpertyDict)


class DataGridMethod(uiautomation.DataGridControl, Method):
    def __init__(self, element=0, searchFromControl=None, searchDepth=0xFFFFFFFF, searchWaitTime=0.5, foundIndex=1,
                 **searchPorpertyDict):
        uiautomation.DataGridControl.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex,
                                              **searchPorpertyDict)

    def GetListCount(self):
        if not self.Exists():
            print("DataGrid Control Not Found, Please Double Check the Parameter!!")
            return
        return self.CurrentRowCount()

    def Export(self, filename):
        """
        This class has default method 'getitem', and this method has a issue:
        when the scrollbar need to move for more than 4 or 5 times, the element will no be found anymore
        before get item please maximum the window
        """
        rows = self.GetListCount()
        cols = self.CurrentColumnCount()
        for i in range(rows):
            for j in range(cols):
                element = self.GetItem(i, j)
                if element == None:
                    print(element, "not found")
                    continue
                name = element.Name
                '''
                判断当为custom类型并包含button或者checkbox 类型的时候
                '''
                if element.ControlType == uiautomation.ControlType.ButtonControl:
                    if element.IsTogglePatternAvailable():
                        name = element.CurrentToggleState()
                    else:
                        name = element.ClassName
                elif element.ControlType == uiautomation.ControlType.CheckBoxControl:
                    name = element.CurrentToggleState()
                    if name:
                        name = "Enable"
                    else:
                        name = "Disable"
                elif element.ControlType == uiautomation.ControlType.CustomControl:
                    if element.ButtonControl(foundIndex=1).Exists(maxSearchSeconds=1):
                        subButton = element.ButtonControl(foundIndex=1)
                        if subButton.IsTogglePatternAvailable():
                            name = subButton.CurrentToggleState()
                            if name:
                                name = "Enable"
                            else:
                                name = "Disable"
                        else:
                            name = subButton.ClassName
                    if element.CheckBoxControl(foundIndex=1).Exists(maxSearchSeconds=1):
                        subButton = element.CheckBoxControl(foundIndex=1)
                        name = subButton.CurrentToggleState()
                        if name:
                            name = "Enable"
                        else:
                            name = "Disable"
                with open('{}.csv'.format(filename), 'a') as f:
                    f.write(str(name))
                    f.write(",")
            with open('{}.csv'.format(filename), 'a') as f:
                f.write("\n")


class RadioButtonMethod(uiautomation.RadioButtonControl, Method):
    def __init__(self, element=0, searchFromControl=None, searchDepth=0xFFFFFFFF, searchWaitTime=0.5, foundIndex=1,
                 **searchPorpertyDict):
        uiautomation.RadioButtonControl.__init__(self, element, searchFromControl, searchDepth, searchWaitTime,
                                                 foundIndex, **searchPorpertyDict)


class ScrollBarMethod(uiautomation.ScrollBarControl, Method):
    def __init__(self, element=0, searchFromControl=None, searchDepth=0xFFFFFFFF, searchWaitTime=0.5, foundIndex=1,
                 **searchPorpertyDict):
        uiautomation.ScrollBarControl.__init__(self, element, searchFromControl, searchDepth, searchWaitTime,
                                               foundIndex,
                                               **searchPorpertyDict)


class TreeMethod(uiautomation.TreeControl, Method):
    def __init__(self, element=0, searchFromControl=None, searchDepth=0xFFFFFFFF, searchWaitTime=0.5, foundIndex=1,
                 **searchPorpertyDict):
        uiautomation.TreeControl.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex,
                                          **searchPorpertyDict)


class TreeItemMethod(uiautomation.TreeItemControl, Method):
    def __init__(self, element=0, searchFromControl=None, searchDepth=0xFFFFFFFF, searchWaitTime=0.5, foundIndex=1,
                 **searchPorpertyDict):
        uiautomation.TreeItemControl.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex,
                                              **searchPorpertyDict)


class WindowMethod(uiautomation.WindowControl, Method):
    def __init__(self, element=0, searchFromControl=None, searchDepth=0xFFFFFFFF, searchWaitTime=1, foundIndex=1,
                 **searchPorpertyDict):
        uiautomation.WindowControl.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex,
                                            **searchPorpertyDict)
# QAutils.LaunchAppFromControl("HP Easy Shell Configuration")

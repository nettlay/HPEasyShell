import math
import winreg
import cv2
import numpy as np
import socket
from uiautomation import *
import os
import time
import win32api
import win32con
import subprocess
import platform
import ruamel.yaml as yaml

'''
Some element might mistake recognize automationId or Name:
Minimize: Only Name is useful, AutomationId actually is "" even though we can get the value
'''


class QAUtils:
    @staticmethod
    def get_window_size():
        # actual size, value will be different with different scaling
        width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
        height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
        return width, height

    @staticmethod
    def get_resolution():
        # if global import pyautogui, Control 'get_window_size will not get the actual window size
        # and will get actual resolution the same with below
        import pyautogui
        x = pyautogui.size()[0]
        y = pyautogui.size()[1]
        return x, y

    @staticmethod
    def get_os():
        return platform.platform()[:10].replace('-', '')

    @staticmethod
    def getNetInfo():
        # ip_pattern = r'(([01]{0,1}\d{0,1}\d|2[0-4]\d|25[0-5])\.){3}(?:[01]{0,1}\d{0,1}\d|2[0-4]\d|25[0-5])'
        name = socket.gethostname()
        ip_addr = socket.gethostbyname(name)
        rs = os.popen('ipconfig /all').read()
        rs_list = re.search('Ethernet adapter(.*?)Subnet Mask', rs, re.S).group().split('\n')
        for line in rs_list:
            if "Physical Address" in line:
                mac_addr = line.split(':')[1].strip().replace('-', ':')
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
        month = int(time.strftime('%m'))
        day = int(time.strftime('%d'))
        if '%Y' in fmt:
            year = int(time.strftime('%Y'))
        elif '%y' in fmt:
            year = int(time.strftime('%y'))
        else:
            year = ''
        #  force convert fmt to time as given format)
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
        for si in sidInfo:
            if si == "":
                continue
            ls = si.split(' S')
            if ls[0].strip() == 'Name':
                continue
            userList[ls[0].strip()] = 'S{}'.format(ls[1]).strip()
        return userList

    @staticmethod
    def Wait(waitTime=3):
        time.sleep(waitTime)

    @staticmethod
    def SendKeys(keys, waitTime=0.1):
        return SendKeys(keys, waitTime=waitTime)

    @staticmethod
    def SendKey(key, waitTime=0.1, count=1):
        for t in range(count):
            SendKey(key, waitTime=waitTime)

    @staticmethod
    def ControlFromCursor():
        return ControlFromCursor()

    @staticmethod
    def GetWindow(name):
        # Special for force convert to window Control
        time.sleep(3)
        if Control(RegexName=name).Exists(0, 0):
            return WindowControl(RegexName=name)
        else:
            return False

    @staticmethod
    def Reboot(wait_reboot=3, wait_time=20):
        os.system('shutdown -r -t {}'.format(wait_reboot))
        time.sleep(wait_time)

    @staticmethod
    def SwitchUser(username="Admin", password="Admin", domain=""):
        """
        Before program, Please Add HKLM\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon
        To Registry exclusions of Write Filter
        """
        root = win32con.HKEY_LOCAL_MACHINE
        key = win32api.RegOpenKeyEx(root, 'SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon', 0,
                                    win32con.KEY_ALL_ACCESS | win32con.KEY_WOW64_64KEY | win32con.KEY_WRITE)
        win32api.RegSetValueEx(key, "DefaultUserName", 0, win32con.REG_SZ, username)
        win32api.RegSetValueEx(key, "DefaultPassWord", 0, win32con.REG_SZ, password)
        if domain != "":
            win32api.RegSetValueEx(key, "DefaultDomain", 0, win32con.REG_SZ, domain)

    @staticmethod
    def Write2File(filename, src):
        TxtUtils(filename, 'a').set_msg(src)

    @staticmethod
    def LaunchAppFromControl(name):
        """
        Start application via Control panel
        :param name: app name shown in Control panel
        :return: None
        """
        os.system("Control panel")
        mainWnd = ""
        for t in range(10):
            print(t)
            if not WindowControl(Name="All Control Panel Items").Exists(0, 0):
                print("not exist")
                continue
            else:
                mainWnd = WindowControl(Name="All Control Panel Items")
                break

        mainWnd.Maximize()
        ButtonControl(AutomationId="ActionButton").Click()
        Control(AutomationId="51").Click()
        ControlPane = mainWnd.PaneControl(AutomationId='CategoryPanel')
        items = ControlPane.GetChildren()
        for item in items:
            if item.Name == name:
                item.Click()
                time.sleep(2)
                break
        if WindowControl(Name="All Control Panel Items").Exists():
            WindowControl(Name="All Control Panel Items").GetWindowPattern().Close()

    @staticmethod
    def GetInstallationList(filename):
        """
        Get installation application in program and features, write to filename
        :param filename: stored filename
        :return: None
        """
        os.system('Control panel')
        mainWnd = WindowControl(Name="All Control Panel Items")
        mainWnd.Maximize()
        ControlPane = mainWnd.PaneControl(AutomationId='CategoryPanel')
        items = ControlPane.GetChildren()
        for item in items:
            if item.Name == "Programs and Features":
                item.Click()
        redirectWnd = WindowControl(Name="Programs and Features")
        installationPool = redirectWnd.ListControl(AutomationId="1")
        items = installationPool.GetChildren()
        for item in items:
            if item.ControlType == ControlType.ScrollBarControl:
                continue
            subitems = item.GetChildren()
            QAUtils.Write2File(filename, "{}:{}\n".format(subitems[0].Name, subitems[4].Name))
        redirectWnd.GetWindowPattern().Close()

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
        path = '{}\\{}'.format(QAUtils.getSid()[os.getlogin()], path)
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
    def getPicRGB(file):
        img1 = cv2.imread(file)
        im_B_mean = np.mean(img1[:, :, 0])
        im_G_mean = np.mean(img1[:, :, 1])
        im_R_mean = np.mean(img1[:, :, 2])
        return [im_R_mean, im_G_mean, im_B_mean]

    @staticmethod
    def compareByRGB(rgb1, rgb2):
        r = (rgb1[0] - rgb2[0]) / 256
        g = (rgb1[1] - rgb2[1]) / 256
        b = (rgb1[2] - rgb2[2]) / 256
        return 1 - math.sqrt(r * r + g * g + b * b)

    @staticmethod
    def AddProxy(user, server='15.85.199.199:8080', override='*.sh.dto;15.83.*.*'):
        """
         User must be actual user that exist in windows
         server: '15.85.199.199:8080'
         override: '*.sh.dto;15.83.*.*'
        """
        sidDict = QAUtils.getSid()
        sid = sidDict[user]
        root = winreg.HKEY_USERS
        path = r"{}\Software\Microsoft\Windows\CurrentVersion\Internet Settings".format(sid)
        key = winreg.OpenKeyEx(root, path, access=winreg.KEY_ALL_ACCESS | winreg.KEY_WOW64_64KEY)
        winreg.SetValueEx(key, 'ProxyServer', 0, winreg.REG_SZ, server)
        winreg.SetValueEx(key, 'ProxyOverride', 0, winreg.REG_SZ, override)
        winreg.SetValueEx(key, 'ProxyEnable', 0, winreg.REG_DWORD, 1)


class YmlUtils:
    def __init__(self, filename):
        with open(filename) as f:
            self._content = yaml.safe_load(f)

    def get_item(self):
        return self._content

    def get_item_keys(self):
        return self._content.keys()

    def get_sub_item(self, item):  # the same with sub item value
        return self._content[item]

    def get_sub_item_keys(self, item):
        pass

    def get_sub_item_value(self, item, sub_item_key):
        pass


class TxtUtils:
    def __init__(self, filename, mode='r', encoding='utf8'):
        self._filename = filename
        self._mode = mode
        self._encoding = encoding

    def _get_read(self):
        return open(self._filename, self._mode, encoding=self._encoding)

    def _get_write(self):
        return open(self._filename, self._mode, encoding=self._encoding)

    def get_lines(self):
        f = self._get_read()
        lines = f.readlines()
        f.close()
        return lines

    def get_source(self):
        f = self._get_read()
        source = f.read()
        f.close()
        return source

    def set_msg(self, msg):
        f = self._get_read()
        f.write(msg)
        f.close()

    def replace_msg(self, new, old):
        data = self.get_source()
        new_data = data.replace(old, new)
        self.set_msg(new_data)


class ButtonControl(Control):
    def __init__(self, searchFromControl: Control = None, searchDepth: int = 0xFFFFFFFF,
                 searchWaitTime: float = SEARCH_INTERVAL, foundIndex: int = 1, element=None, **searchProperties):
        Control.__init__(self, searchFromControl, searchDepth, searchWaitTime, foundIndex, element, **searchProperties)
        self.AddSearchProperties(ControlType=ControlType.ButtonControl)

    def GetExpandCollapsePattern(self) -> ExpandCollapsePattern:
        """
        Return `ExpandCollapsePattern` if it supports the pattern else None(Conditional support according to MSDN).
        """
        return self.GetPattern(PatternId.ExpandCollapsePattern)

    def GetInvokePattern(self) -> InvokePattern:
        """
        Return `InvokePattern` if it supports the pattern else None(Conditional support according to MSDN).
        """
        return self.GetPattern(PatternId.InvokePattern)

    def GetTogglePattern(self) -> TogglePattern:
        """
        Return `TogglePattern` if it supports the pattern else None(Conditional support according to MSDN).
        """
        return self.GetPattern(PatternId.TogglePattern)

    def Active(self, waitTime=0.1):
        """
        Toggle or Invoke button with build-in Control instead of Click
        """
        if self.GetTogglePattern() is not None:
            self.SetFocus()
            self.GetTogglePattern().Toggle(waitTime)
        elif self.GetInvokePattern() is not None:
            self.SetFocus()
            self.GetInvokePattern().Invoke(waitTime)
        else:
            self.SetFocus()

    def GetStatus(self):
        if not self.Exists(1, 1):
            print("Button not Found, Please Double check your parameter!!")
            return None
        if self.GetTogglePattern() is not None:
            return self.GetTogglePattern().ToggleState
        else:
            return None

    def Enable(self):
        if not self.Exists(1, 1):
            print("Button not Found, Please Double check your parameter!!")
            return
        toggle_pattern = self.GetTogglePattern()
        if toggle_pattern is not None:
            if toggle_pattern.ToggleState:
                return
            else:
            # sometimes toggle method will not affect ui except clicking
            # but sometimes button is offscreen, so firstly setfocus
                if self.IsOffscreen:
                    toggle_pattern.Toggle()
                else:
                    self.SetFocus()
                    self.Click()
        else:
            print('Button do not support Enable Control')

    def Disable(self):
        if not self.Exists(1, 1):
            print("Button not Found, Please Double check your parameter!!")
            return
        toggle_pattern = self.GetTogglePattern()
        if toggle_pattern is not None:
            if toggle_pattern.ToggleState:
                # sometimes toggle method will not affect ui except clicking
                # but sometimes button is offscreen, so firstly setfocus
                if self.IsOffscreen:
                    toggle_pattern.Toggle()
                else:
                    self.SetFocus()
                    self.Click()
            else:
                return
        else:
            print('Button do not support Disable Control')


class EditControl(Control):
    def __init__(self, searchFromControl: Control = None, searchDepth: int = 0xFFFFFFFF, searchWaitTime: float = SEARCH_INTERVAL, foundIndex: int = 1, element=None, **searchProperties):
        Control.__init__(self, searchFromControl, searchDepth, searchWaitTime, foundIndex, element, **searchProperties)
        self.AddSearchProperties(ControlType=ControlType.EditControl)

    def GetRangeValuePattern(self) -> RangeValuePattern:
        """
        Return `RangeValuePattern` if it supports the pattern else None(Conditional support according to MSDN).
        """
        return self.GetPattern(PatternId.RangeValuePattern)

    def GetTextPattern(self) -> TextPattern:
        """
        Return `TextPattern` if it supports the pattern else None(Conditional support according to MSDN).
        """
        return self.GetPattern(PatternId.TextPattern)

    def GetValuePattern(self) -> ValuePattern:
        """
        Return `ValuePattern` if it supports the pattern else None(Conditional support according to MSDN).
        """
        return self.GetPattern(PatternId.ValuePattern)

    def SetValue(self, value):
        self.GetValuePattern().SetValue(value)

    def GetValue(self):
        return self.GetValuePattern().Value


class TextControl(Control):
    def __init__(self, searchFromControl: Control = None, searchDepth: int = 0xFFFFFFFF,
                 searchWaitTime: float = SEARCH_INTERVAL, foundIndex: int = 1, element=None, **searchProperties):
        Control.__init__(self, searchFromControl, searchDepth, searchWaitTime, foundIndex, element, **searchProperties)
        self.AddSearchProperties(ControlType=ControlType.TextControl)

    def GetGridItemPattern(self) -> GridItemPattern:
        """
        Return `GridItemPattern` if it supports the pattern else None(Conditional support according to MSDN).
        """
        return self.GetPattern(PatternId.GridItemPattern)

    def GetTableItemPattern(self) -> TableItemPattern:
        """
        Return `TableItemPattern` if it supports the pattern else None(Conditional support according to MSDN).
        """
        return self.GetPattern(PatternId.TableItemPattern)

    def GetTextPattern(self) -> TextPattern:
        """
        Return `TextPattern` if it supports the pattern else None(Conditional support according to MSDN).
        """
        return self.GetPattern(PatternId.TextPattern)

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
        if str(obj).upper() in self.Name.upper():
            return True
        else:
            return False


def getElementByType(controlType, **kwargs):
    if controlType.upper() == "HYPERLINK":
        return HyperlinkControl(**kwargs)
    elif controlType.upper() == 'IMAGE':
        ImageControl(**kwargs)
    elif controlType.upper() == 'APPBAR':
        AppBarControl(**kwargs)
    elif controlType.upper() == 'CALENDAR':
        CalendarControl(**kwargs)
    elif controlType.upper() == 'CUSTOM':
        CustomControl(**kwargs)
    elif controlType.upper() == 'DOCUMENT':
        DocumentControl(**kwargs)
    elif controlType.upper() == 'GROUP':
        GroupControl(**kwargs)
    elif controlType.upper() == 'HEADER':
        HeaderControl(**kwargs)
    elif controlType.upper() == 'HEADERITEM':
        HeaderItemControl(**kwargs)
    elif controlType.upper() == 'SPINNER':
        SpinnerControl(**kwargs)
    elif controlType.upper() == 'SLIDER':
        SliderControl(**kwargs)
    elif controlType.upper() == 'SEPARATOR':
        SeparatorControl(**kwargs)
    elif controlType.upper() == 'SEMANTICZOOM':
        SemanticZoomControl(**kwargs)
    elif controlType.upper() == 'MENUITEM':
        MenuItemControl(**kwargs)
    elif controlType.upper() == 'MENU':
        MenuControl(**kwargs)
    elif controlType.upper() == 'MENUBAR':
        MenuBarControl(**kwargs)
    elif controlType.upper() == 'TABLE':
        TableControl(**kwargs)
    elif controlType.upper() == 'STATUSBAR':
        StatusBarControl(**kwargs)
    elif controlType.upper() == 'SPLITBUTTON':
        SplitButtonControl(**kwargs)
    elif controlType.upper() == 'THUMB':
        ThumbControl(**kwargs)
    elif controlType.upper() == 'TITLEBAR':
        TitleBarControl(**kwargs)
    elif controlType.upper() == 'TOOLTIP':
        return ToolTipControl(**kwargs)
    elif controlType.upper() == "BUTTON":
        return ButtonControl(**kwargs)
    elif controlType.upper() == "COMBOX":
        return ComboBoxControl(**kwargs)
    elif controlType.upper() == "CHECKBOX":
        return CheckBoxControl(**kwargs)
    elif controlType.upper() == "DATAGRID":
        return DataGridControl(**kwargs)
    elif controlType.upper() == "DATAITEM":
        return DataItemControl(**kwargs)
    elif controlType.upper() == "EDIT":
        return EditControl(**kwargs)
    elif controlType.upper() == "LIST":
        return ListControl(**kwargs)
    elif controlType.upper() == "LISTITEM":
        return ListItemControl(**kwargs)
    elif controlType.upper() == "PANE":
        return PaneControl(**kwargs)
    elif controlType.upper() == "RADIOBUTTON":
        return RadioButtonControl(**kwargs)
    elif controlType.upper() == "SCROLLBAR":
        return ScrollBarControl(**kwargs)
    elif controlType.upper() == "TEXT":
        return TextControl(**kwargs)
    elif controlType.upper() == "TREE":
        return TreeControl(**kwargs)
    elif controlType.upper() == "TREEITEM":
        return TreeItemControl(**kwargs)
    elif controlType.upper() == "WINDOW":
        return WindowControl(**kwargs)
    elif controlType.upper() == 'PROGRESSBAR':
        return ProgressBarControl(**kwargs)
    elif controlType.upper() == "TAB":
        return TabControl(**kwargs)
    elif controlType.upper() == "TABITEM":
        return TabItemControl(**kwargs)
    else:
        return Control(**kwargs)

import math
import winreg
import cv2
import numpy as np
import socket
import re
import uiautomation
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
        # if global import pyautogui, method 'get_window_size will not get the actual window size
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
        for i in sidInfo:
            if i == "":
                continue
            ls = i.split(' S')
            if ls[0].strip() == 'Name':
                continue
            userList[ls[0].strip()] = 'S{}'.format(ls[1]).strip()
        return userList

    @staticmethod
    def Wait(waitTime=3):
        time.sleep(waitTime)

    @staticmethod
    def SendKeys(keys, waitTime=0.1):
        return uiautomation.SendKeys(keys, waitTime=waitTime)

    @staticmethod
    def SendKey(key, waitTime=0.1, count=1):
        for i in range(count):
            uiautomation.SendKey(key, waitTime=waitTime)

    @staticmethod
    def Keys():
        return uiautomation.Keys

    @staticmethod
    def ControlFromCursor():
        return uiautomation.ControlFromCursor()

    @staticmethod
    def GetWindow(name):
        # Special for force convert to window method
        time.sleep(3)
        if Method(RegexName=name).Exists(0, 0):
            return WindowMethod(RegexName=name)
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
        :param name: app name shown in control panel
        :return: None
        """
        os.system("control panel")

        for i in range(10):
            print(i)
            if not WindowMethod(Name="All Control Panel Items").Exists(0, 0):
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
                item.Click()
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
                item.Click()
        redirectWnd = WindowMethod(Name="Programs and Features")
        installationPool = redirectWnd.ListControl(AutomationId="1")
        items = installationPool.GetChildren()
        for item in items:
            if item.ControlType == uiautomation.ControlType.ScrollBarControl:
                continue
            subitems = item.GetChildren()
            QAUtils.Write2File(filename, "{}:{}\n".format(subitems[0].Name, subitems[4].Name))
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


"""
    # designed for compare pictures' different, but do not support pure color
    # change to another function to compare this
        def __make_regalur_image(self, img, size=(256, 256)):
            return img.resize(size).convert('RGB')
    
        # 几何转变，全部转化为256*256像素大小
    
        def __split_image(self, img, part_size=(64, 64)):
            w, h = img.size
            pw, ph = part_size
            assert w % pw == h % ph == 0
    
            return [img.crop((i, j, i + pw, j + ph)).copy() for i in range(0, w, pw) for j in range(0, h, ph)]
    
        # region = img.crop(box)
        # 将img表示的图片对象拷贝到region中，这个region可以用来后续的操作（region其实就是一个
        # image对象，box是个四元组（上下左右））
    
        def __hist_similar(self, lh, rh):
            assert len(lh) == len(rh)
            return sum(1 - (0 if l == r else float(abs(l - r)) / max(l, r)) for l, r in zip(lh, rh)) / len(lh)
    
        # 好像是根据图片的左右间隔来计算某个长度，zip是可以接受多个x,y,z数组值统一输出的输出语句
        def __calc_similar(self, li, ri):
            # return hist_similar(li.histogram(), ri.histogram())
            return sum(
                self.__hist_similar(l.histogram(), r.histogram()) for l, r in
                zip(self.__split_image(li), self.__split_image(ri))) / 16.0  # 256>64
    
        # 其中histogram()对数组x（数组是随机取样得到的）进行直方图统计，它将数组x的取值范围分为100个区间，
        # 并统计x中的每个值落入各个区间中的次数。histogram()返回两个数组p和t2，
        # 其中p表示各个区间的取样值出现的频数，t2表示区间。
        # 大概是计算一个像素点有多少颜色分布的
        # 把split_image处理的东西zip一下，进行histogram,然后得到这个值
    
        def calc_similar_by_path(self, lf, rf):
            li, ri = self.__make_regalur_image(Image.open(lf)), self.__make_regalur_image(Image.open(rf))
            return self.__calc_similar(li, ri)
    """


class Method(uiautomation.Control):
    def __init__(self, element=0, searchFromControl=None, searchDepth=0xFFFFFFFF, searchWaitTime=0.5, foundIndex=1,
                 **searchPorpertyDict):
        uiautomation.Control.__init__(self, element, searchFromControl, searchDepth,
                                      searchWaitTime, foundIndex, **searchPorpertyDict)
        """
        element: int
        searchFromControl: Control, if is None, search from root control(Desktop)
        searchDepth: int, max search depth from searchFromControl
        foundIndex: int, value must >= 1
        searchWaitTime: float, wait searchWaitTime before every search
        searchPorpertyDict: a dict that defines how to search, the following keys can be used
                            ControlType: int, a value in class ControlType
                            ClassName: str or unicode
                            AutomationId: str or unicode
                            Name: str or unicode
                            SubName: str or unicode
                            RegexName: str or unicode, supports regex
                            Depth: int, relative depth from searchFromControl, if set, searchDepth will be set to Depth too
                            Compare: custom compare function(control, depth) returns a bool value
                            others:[LocalizedControlType,ClassName,ProcessId...]
        """

    def Drag(self, x=0, y=0):
        x1, y1 = self.BoundingRectangle[0], self.BoundingRectangle[1]
        print(x1, y1)
        uiautomation.DragDrop(x1, y1, x1 + x, y1 + y)

    def IsShown(self):
        if self.IsOffScreen:
            return False
        else:
            return True

    def waitExists(self, wait_cycles=1):
        for i in range(wait_cycles):
            if self.Exists(0, 0):
                return self
            else:
                time.sleep(1)
                continue
        return None


class ButtonMethod(uiautomation.ButtonControl, Method):
    def __init__(self, element=0, searchFromControl=None, searchDepth=0xFFFFFFFF, searchWaitTime=0.5, foundIndex=1,
                 **searchPorpertyDict):
        uiautomation.ButtonControl.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex,
                                            **searchPorpertyDict)

    def Active(self, waitTime=0.1):
        """
        Toggle or Invoke button with build-in method instead of Click
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
        if not self.Exists():
            print("Button not Found, Please Double check your parameter!!")
            return None
        if self.GetTogglePattern() is not None:
            return self.GetTogglePattern().ToggleState
        else:
            return None

    def Enable(self):
        if not self.Exists():
            print("Button not Found, Please Double check your parameter!!")
            return
        toggle_pattern = self.GetTogglePattern()
        if toggle_pattern is not None:
            if toggle_pattern.ToggleState:
                return
            else:
                self.SetFocus()
                toggle_pattern.Toggle()
        else:
            print('Button do not support Enable method')

    def Disable(self):
        if not self.Exists():
            print("Button not Found, Please Double check your parameter!!")
            return
        toggle_pattern = self.GetTogglePattern()
        if toggle_pattern is not None:
            if toggle_pattern.ToggleState:
                self.SetFocus()
                toggle_pattern.Toggle()
            else:
                return
        else:
            print('Button do not support Disable method')


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
        if str(obj).upper() in self.Name.upper():
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


class ListMethod(uiautomation.ListControl, Method):
    def __init__(self, element=0, searchFromControl=None, searchDepth=0xFFFFFFFF, searchWaitTime=0.5, foundIndex=1,
                 **searchPorpertyDict):
        uiautomation.ListControl.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex,
                                          **searchPorpertyDict)


class ListItemMethod(uiautomation.ListItemControl, Method):
    def __init__(self, element=0, searchFromControl=None, searchDepth=0xFFFFFFFF, searchWaitTime=0.5, foundIndex=1,
                 **searchPorpertyDict):
        uiautomation.ListItemControl.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex,
                                              **searchPorpertyDict)


class ComboBoxMethod(uiautomation.ComboBoxControl, Method):
    def __init__(self, element=0, searchFromControl=None, searchDepth=0xFFFFFFFF, searchWaitTime=0.5, foundIndex=1,
                 **searchPorpertyDict):
        uiautomation.ComboBoxControl.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex,
                                              **searchPorpertyDict)


class TabMethod(uiautomation.TabControl, Method):
    def __init__(self, element=0, searchFromControl=None, searchDepth=0xFFFFFFFF, searchWaitTime=0.5,
                 foundIndex=1, **searchPorpertyDict):
        uiautomation.TabControl.__init__(self, element, searchFromControl, searchDepth, searchWaitTime,
                                         foundIndex,
                                         **searchPorpertyDict)


class TabItemMethod(uiautomation.TabItemControl, Method):
    def __init__(self, element=0, searchFromControl=None, searchDepth=0xFFFFFFFF, searchWaitTime=0.5, foundIndex=1,
                 **searchPorpertyDict):
        uiautomation.TabItemControl.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex,
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
    def __init__(self, searchFromControl= None, searchDepth= 0xFFFFFFFF, searchWaitTime= 1, foundIndex= 1, element =None, **searchProperties):
        uiautomation.WindowControl.__init__(self, searchFromControl,searchDepth, searchWaitTime,  foundIndex, element, **searchProperties)


class ProgressBarMethod(uiautomation.ProgressBarControl, Method):
    def __init__(self, element=0, searchFromControl=None, searchDepth=0xFFFFFFFF, searchWaitTime=1, foundIndex=1,
                 **searchPorpertyDict):
        uiautomation.ProgressBarControl.__init__(self, element, searchFromControl, searchDepth, searchWaitTime,
                                                 foundIndex,
                                                 **searchPorpertyDict)


# QAutils.LaunchAppFromControl("HP Easy Shell Configuration")


def getElementByType(controltype, **kwargs):
    if controltype.upper() in ["APPBAR", "CALENDAR" "CUSTOM", "DOCUMENT", "GROUP", "HEADER",
                               "HEADERITEM", "HYPERLINK", "IMAGE", "MENUBAR", "MENU", "MENUITEM",
                               "SEMANTICZOOM", "SEPARATOR", "SLIDER", "SPINNER", "SPLITBUTTON", "STATUSBAR",
                               "TABLE", "THUMB", "TITLEBAR", "TOOLTIP", "TOOLTIP", ]:
        return Method(**kwargs)
    elif controltype.upper() == "BUTTON":
        return ButtonMethod(**kwargs)
    elif controltype.upper() == "COMBOX":
        return ComboBoxMethod(**kwargs)
    elif controltype.upper() == "CHECKBOX":
        return CheckBoxMethod(**kwargs)
    elif controltype.upper() == "DATAGRID":
        return DataGridMethod(**kwargs)
    elif controltype.upper() == "DATAITEM":
        return DataItemMethod(**kwargs)
    elif controltype.upper() == "EDIT":
        return EditMethod(**kwargs)
    elif controltype.upper() == "LIST":
        return ListMethod(**kwargs)
    elif controltype.upper() == "LISTITEM":
        return ListItemMethod(**kwargs)
    elif controltype.upper() == "PANE":
        return PaneMethod(**kwargs)
    elif controltype.upper() == "RADIOBUTTON":
        return RadioButtonMethod(**kwargs)
    elif controltype.upper() == "SCROLLBAR":
        return ScrollBarMethod(**kwargs)
    elif controltype.upper() == "TEXT":
        return TextMethod(**kwargs)
    elif controltype.upper() == "TREE":
        return TreeMethod(**kwargs)
    elif controltype.upper() == "TREEITEM":
        return TreeItemMethod(**kwargs)
    elif controltype.upper() == "WINDOW":
        return WindowMethod(**kwargs)
    elif controltype.upper() == 'PROGRESSBAR':
        return ProgressBarMethod(**kwargs)
    elif controltype.upper() == "TAB":
        return TabMethod(**kwargs)
    elif controltype.upper() == "TABITEM":
        return TabItemMethod(**kwargs)

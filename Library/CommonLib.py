import ftplib
import uuid
import math
import winreg
import cv2
import numpy as np
import socket
from uiautomation import *
import uiautomation
import os
import time
import win32api
import win32con
import win32net
import win32netcon
import subprocess
import platform
import ruamel.yaml as yaml
import re
'''
Some element might mistake recognize automationId or Name:
Minimize: Only Name is useful, AutomationId actually is "" even though we can get the value
'''


class QAUtils:
    @staticmethod
    def import_cert(file):
        os.system(r'c:\windows\sysnative\certutil -addstore root {}'.format(file))

    @staticmethod
    def get_cert_id(name):
        exist = None
        f = os.popen(r'c:\windows\sysnative\certutil -store root')
        rs = f.read()
        patter = '================ Certificate(.*?)===========.*?Issuer: (.*?)NotBefore: '
        a = re.findall(patter, rs, re.S)
        for i in a:
            if name.upper() in i[1].upper():
                exist = int(i[0].strip())
                break
        return exist

    @staticmethod
    def del_cert(index:int):
        os.system(r'c:\windows\sysnative\certutil -delstore root {}'.format(index))

    @staticmethod
    def hex2rgb(hexcolor):
        rgb = [(hexcolor >> 16) & 0xff,
               (hexcolor >> 8) & 0xff,
               hexcolor & 0xff
               ]
        return rgb

    def rgb2hex(rgbcolor):
        r, g, b = rgbcolor
        return (r << 16) + (g << 8) + b

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

    # @staticmethod
    # def getNetInfo():
    #     # ip_pattern = r'(([01]{0,1}\d{0,1}\d|2[0-4]\d|25[0-5])\.){3}(?:[01]{0,1}\d{0,1}\d|2[0-4]\d|25[0-5])'
    #     name = socket.gethostname()
    #     ip_addr = socket.gethostbyname(name)
    #     rs = os.popen('ipconfig /all').read() # if run in service, there is exception oserror:[winerror 6] the handle is invalid
    #     rs_list = re.search('Ethernet adapter(.*?)Subnet Mask', rs, re.S).group().split('\n')
    #     for line in rs_list:
    #         if "Physical Address" in line:
    #             mac_addr = line.split(':')[1].strip().replace('-', ':')
    #             return {'MAC': mac_addr, 'IP': ip_addr}
    @staticmethod
    def getNetInfo():
        # ip_pattern = r'(([01]{0,1}\d{0,1}\d|2[0-4]\d|25[0-5])\.){3}(?:[01]{0,1}\d{0,1}\d|2[0-4]\d|25[0-5])'
        name = socket.gethostname()
        ip_addr = socket.gethostbyname(name)
        mac_1 = uuid.UUID(int = uuid.getnode()).hex[-12:]
        mac = ":".join([mac_1[e:e+2] for e in range(0,11,2)]).upper()
        return {'MAC': mac, 'IP': ip_addr}

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
            WindowControl(Name="All Control Panel Items").Close()

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
        redirectWnd().Close()

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


class Reg_Utils:
    def __init__(self, root='local'):
        if root.upper() == 'LOCAL':
            self.root = win32con.HKEY_LOCAL_MACHINE
        else:
            self.root = win32con.HKEY_CURRENT_USER
        self.flags = win32con.WRITE_OWNER | win32con.KEY_WOW64_64KEY | win32con.KEY_ALL_ACCESS
        self.reg_type = {0: win32con.REG_SZ, 1:win32con.REG_DWORD, 2:win32con.REG_BINARY}

    def open(self, path):
        key = win32api.RegOpenKeyEx(self.root, path, 0, self.flags)
        return key

    def isKeyExist(self, path):
        try:
            return self.open(path)
        except:
            return False

    def isValueExist(self, key, value):
        try:
            return win32api.RegQueryValueEx(key, value)
        except:
            return False

    def list_key(self, key):
        sub_key_list = []
        if key:
            for sub_key in win32api.RegEnumKeyEx(key):
                # win32api.RegDeleteKeyEx(sub_key)
                sub_key_list.append(sub_key[0])
            return sub_key_list
        else:
            return None

    def list_value(self, key):
        try:
            i = 0
            while 1:
                print(win32api.RegEnumValue(key, i))
                i += 1
        except:
            pass

    def create_key(self, path):
        key, _ = win32api.RegCreateKeyEx(self.root, path, self.flags)
        return key

    def del_key(self, path):
        key = win32api.RegOpenKeyEx(self.root, path, 0, self.flags)
        m_item = win32api.RegEnumKeyEx(key)
        if not m_item:
            reg_parent, subkey_name = os.path.split(path)  # 获得父路径名字 和自己的名字，而不是路径
            try:
                key_parent = win32api.RegOpenKeyEx(self.root, reg_parent, 0, self.flags)  # 看这个节点是否可被访问
                win32api.RegDeleteKeyEx(key_parent, subkey_name)  # 删除这个节点
                return
            except Exception as e:
                print("Bently 被拒绝访问")
                return

        for item in win32api.RegEnumKeyEx(key):  # 递归加子节点
            strRecord = item[0]  # 采用key的第一个节点，item里面是元组，获取第一个名字。就是要的子项名字
            newpath = path + '\\' + strRecord
            self.del_key(newpath)

            # 删除父节点
        root_parent, child_name = os.path.split(path)
        try:  # 看这个节点是否可被访问
            current_parent = win32api.RegOpenKeyEx(self.root, root_parent, 0, self.flags)
            win32api.RegDeleteKeyEx(current_parent, child_name)
        except Exception as e:
            print("Bently 被拒绝访问")
            return

    def clear_subkeys(self, path):
        key = self.isKeyExist(path)
        for subkey in self.list_key(key):
            self.del_key('{}\{}'.format(path, subkey))
        self.close(key)

    def get_value(self, key, value):
        if self.isValueExist(key, value):
            value, type = self.isValueExist(key, value)
            return (value, type)
        else:
            return None

    def create_value(self, key, valueName, regType=0, content=''):
        win32api.RegSetValueEx(key, valueName, 0, self.reg_type[regType], content)

    def del_value(self, key, value):
        try:
            win32api.RegDeleteValue(key, value)
        except:
            print('Value {} not Exist'.format(value))

    def close(self, key):
        win32api.RegCloseKey(key)


class User_Group:
    def __init__(self, user='test', password='test'):
        self.user = user
        self.passwd = password
        self.user_info = dict(
            name=self.user,
            password=self.passwd,
            priv=win32netcon.USER_PRIV_USER,
            home_dir=None,
            comment=None,
            flag=win32netcon.UF_SCRIPT | win32netcon.UF_DONT_EXPIRE_PASSWD | win32netcon.UF_NORMAL_ACCOUNT,
            script_path=None
        )
        self.group_info = dict(
            domainandname=self.user
        )

    def del_user(self):
        try:
            win32net.NetUserDel(None, self.user)
        except:
            pass

    def add_user(self,):
        try:
            win32net.NetUserAdd(None, 1, self.user_info)
        except:
            print('Add new user [{}] fail'.format(self.user_info['name']))
            pass

    def change_passwd(self, new_passwd):
        win32net.NetUserChangePassword(None, self.user, self.passwd, new_passwd)

    def add_user_to_group(self, group='Administrators'):
        try:
            win32net.NetLocalGroupAddMembers(None, group, 3, [self.group_info])
        except:
            pass


class YmlUtils:
    def __init__(self, filename):
        self.__fiename = filename
        if not os.path.exists(self.__fiename):
            with open(self.__fiename, 'w', encoding='utf-8') as f:
                f.close()
        self._content = self.get_item()

    def get_item(self):
        with open(self.__fiename, encoding='utf-8') as f:
            return yaml.safe_load(f)

    def get_item_keys(self):
        return self._content.keys()

    def get_sub_item(self, item):  # the same with sub item value
        return self._content[item]

    def get_sub_item_keys(self, item):
        pass

    def get_sub_item_value(self, item, sub_item_key):
        pass

    def write(self, msg):
        with open(self.__fiename, 'w', encoding='utf-8') as f:
            yaml.safe_dump(msg, f)
            f.flush()
            f.close()
# yml = YmlUtils("Z:\\workspace3\\hpeasyshell\\test_report\\result.yml")
# case = yml.get_item()
# if not case:
#     yml.write([{'case_name': 'test2', 'result': '', 'steps': [], 'uut_name': '145666'}])
#     time.sleep(9)
#     case = yml.get_item()
#     print(case)
# for index in range(len(case)):
#     time.sleep(1)
#     if case[index].get('case_name') == 'test4':
#         print('if')
#         case1 = case[index]
#         case1['steps'].append({"test":"2323", "step_name":"step_n"})
#     else:
#         print('else')
#         case.extend([{'case_name': 'test2', 'result': '', 'steps': [], 'uut_name': '145666'}])
# yml.write(case)


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


class FTPUtils:
    def __init__(self, server, username, password):
        self.ftp = ftplib.FTP(server)
        self.ftp.login(username, password)
        self.download_buffer = 1024
        self.upload_buffer = 1024

    def change_dir(self, work_dir):
        self.ftp.cwd(work_dir)

    def get_working_dir(self):
        return self.ftp.pwd()

    def get_item_list(self, work_dir):
        self.ftp.cwd(work_dir)
        return self.ftp.nlst()

    def is_item_file(self, item):
        try:
            self.ftp.cwd(item)
            self.ftp.cwd('..')
            return False
        except ftplib.error_perm as fe:
            if not fe.args[0].startswith('550'):
                raise
            return True

    def download_file(self, file_name, save_as_name):
        file_handler = open(save_as_name, 'wb')
        self.ftp.retrbinary("RETR " + file_name, file_handler.write, self.download_buffer)
        file_handler.close()
        return save_as_name

    def download_dir(self, dir_name, save_as_dir):
        if not os.path.exists(save_as_dir):
            os.mkdir(save_as_dir)
        self.ftp.cwd(dir_name)
        for item in self.ftp.nlst():
            if self.is_item_file(item):
                self.download_file(item, os.path.join(save_as_dir, item))
            else:
                self.download_dir(item, os.path.join(save_as_dir, item))
        self.ftp.cwd("..")
        return save_as_dir

    def upload_file(self, file_name, save_as_name):
        self.ftp.storbinary('STOR ' + save_as_name, open(file_name, 'rb'), self.upload_buffer)
        return save_as_name

    def new_dir(self, dir_name):
        try:
            self.ftp.mkd(dir_name)
        # ignore "directory already exists"
        except ftplib.error_perm as fe:
            if not fe.args[0].startswith('550'):
                raise
        return dir_name

    def upload_dir(self, dir_path, save_as_dir):
        try:
            self.ftp.cwd(save_as_dir)
        except ftplib.error_perm as fe:
            if fe.args[0].startswith('550'):
                self.new_dir(save_as_dir)
                self.ftp.cwd(save_as_dir)
        for item_name in os.listdir(dir_path):
            local_path = os.path.join(dir_path, item_name)
            if os.path.isfile(local_path):
                self.ftp.storbinary('STOR ' + item_name, open(local_path, 'rb'), self.upload_buffer)
            elif os.path.isdir(local_path):
                ftp_folder = self.new_dir(item_name)
                self.upload_dir(local_path, ftp_folder)
                self.ftp.cwd("..")
        return save_as_dir

    def delete_file(self, file_name):
        self.ftp.delete(file_name)
        return file_name

    def delete_dir(self, dir_name):
        # Handle dir_name doesn't exist
        try:
            self.change_dir(dir_name)
            items = self.ftp.nlst()
        except ftplib.all_errors as e:
            print(e)
            return
        for item in items:
            if os.path.split(item)[1] in ('.', '..'):
                continue
            try:
                self.change_dir(item)
                self.change_dir('..')
                self.delete_dir(item)
            except ftplib.all_errors:
                self.delete_file(item)
        try:
            self.change_dir('..')
            self.ftp.rmd(dir_name)
        except ftplib.all_errors as e:
            print(e)
        return dir_name

    def close(self):
        self.ftp.close()


class Control(uiautomation.Control):
    def __init__(self, element, searchFromControl, searchDepth, searchWaitTime,
                 foundIndex, **searchPorpertyDict):
        uiautomation.Control.__init__(self, element=0, searchFromControl=None, searchDepth=0xFFFFFFFF, searchWaitTime=SEARCH_INTERVAL,
                 foundIndex=1, **searchPorpertyDict)
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
        """
        self._element = element
        self._elementDirectAssign = True if element else False
        self.searchFromControl = searchFromControl
        self.searchDepth = searchPorpertyDict.get('Depth', searchDepth)
        self.searchWaitTime = searchWaitTime
        self.foundIndex = foundIndex
        self.searchPorpertyDict = searchPorpertyDict
        regName = searchPorpertyDict.get('RegexName', '')
        self.regexName = re.compile(regName) if regName else None

    def Click(self, ratioX=0.5, ratioY=0.5, simulateMove=False, waitTime=0.1):
        """
        ratioX: float or int, if is int, click left + ratioX, if < 0, click right + ratioX
        ratioY: float or int, if is int, click top + ratioY, if < 0, click bottom + ratioY
        simulateMove: bool, if True, first move cursor to control smoothly
        waitTime: float
        Click(0.5, 0.5): click center
        Click(10, 10): click left+10, top+10
        Click(-10, -10): click right-10, bottom-10
        """
        self.SetFocus()
        x, y = self.MoveCursor(ratioX, ratioY, simulateMove)
        Win32API.MouseClick(x, y, waitTime)


class ButtonControl(Control, ExpandCollapsePattern, InvokePattern, TogglePattern):
    def __init__(self, element=0, searchFromControl=None, searchDepth=0xFFFFFFFF, searchWaitTime=SEARCH_INTERVAL, foundIndex=1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType=ControlType.ButtonControl)

    def Active(self, waitTime=0.1):
        """
        Toggle or Invoke button with build-in Control instead of Click
        """
        if self.IsTogglePatternAvailable():
            self.SetFocus()
            self.Toggle()
        elif self.IsInvokePatternAvailable() is not None:
            self.SetFocus()
            self.Invoke(waitTime)
        else:
            self.SetFocus()

    def GetStatus(self):
        if not self.Exists(1, 1):
            print("Button not Found, Please Double check your parameter!!")
            return None
        if self.IsTogglePatternAvailable():
            return self.CurrentToggleState()
        else:
            return None

    def Enable(self):
        if not self.Exists(1, 1):
            print("Button not Found, Please Double check your parameter!!")
            return
        if self.IsTogglePatternAvailable():
            if self.CurrentToggleState():
                return
            else:
            # sometimes toggle method will not affect ui except clicking
            # but sometimes button is offscreen, so firstly setfocus
                self.SetFocus()
                self.Refind(maxSearchSeconds=TIME_OUT_SECOND, searchIntervalSeconds=self.searchWaitTime)
                if not self.IsOffScreen and self.IsEnabled:
                    self.Click()
        else:
            print('Button do not support Enable Control or is not shown')

    def Disable(self):
        if not self.Exists(1, 1):
            print("Button not Found, Please Double check your parameter!!")
            return
        if self.IsTogglePatternAvailable():
            if self.CurrentToggleState():
                # sometimes toggle method will not affect ui except clicking
                # but sometimes button is offscreen, so firstly setfocus
                self.SetFocus()
                self.Refind(maxSearchSeconds=TIME_OUT_SECOND, searchIntervalSeconds=self.searchWaitTime)
                if not self.IsOffScreen and self.IsEnabled:
                    self.Click()
            else:
                return
        else:
            print('Button do not support Enable Control or is not shown')


class TextControl(Control, GridItemPattern, TableItemPattern, TextPattern):
    def __init__(self, element=0, searchFromControl=None, searchDepth=0xFFFFFFFF, searchWaitTime=SEARCH_INTERVAL, foundIndex=1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType=ControlType.TextControl)

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


def SendKey(key, waitTime=0.1, count=1):
    for t in range(count):
        uiautomation.SendKey(key, waitTime=waitTime)


def combination_key(self, *args):
    # args = [key1, key2, key3, ...]
    for key in args:
        self.KeyDown(key)
    for key in args:
        self.KeyUp(key)


def getElementByType(controlType, **kwargs):
    # parent: Parent element
    # print(controlType, **kwargs)
    if controlType.upper() == "HYPERLINK":
        return HyperlinkControl(**kwargs)
    elif controlType.upper() == 'IMAGE':
        return ImageControl(**kwargs)
    elif controlType.upper() == 'APPBAR':
        return AppBarControl(**kwargs)
    elif controlType.upper() == 'CALENDAR':
        return CalendarControl(**kwargs)
    elif controlType.upper() == 'CUSTOM':
        return CustomControl(**kwargs)
    elif controlType.upper() == 'DOCUMENT':
        return DocumentControl(**kwargs)
    elif controlType.upper() == 'GROUP':
        return GroupControl(**kwargs)
    elif controlType.upper() == 'HEADER':
        return HeaderControl(**kwargs)
    elif controlType.upper() == 'HEADERITEM':
        return HeaderItemControl(**kwargs)
    elif controlType.upper() == 'SPINNER':
        return SpinnerControl(**kwargs)
    elif controlType.upper() == 'SLIDER':
        return SliderControl(**kwargs)
    elif controlType.upper() == 'SEPARATOR':
        return SeparatorControl(**kwargs)
    elif controlType.upper() == 'SEMANTICZOOM':
        return SemanticZoomControl(**kwargs)
    elif controlType.upper() == 'MENUITEM':
        return MenuItemControl(**kwargs)
    elif controlType.upper() == 'MENU':
        return MenuControl(**kwargs)
    elif controlType.upper() == 'MENUBAR':
        return MenuBarControl(**kwargs)
    elif controlType.upper() == 'TABLE':
        return TableControl(**kwargs)
    elif controlType.upper() == 'STATUSBAR':
        return StatusBarControl(**kwargs)
    elif controlType.upper() == 'SPLITBUTTON':
        return SplitButtonControl(**kwargs)
    elif controlType.upper() == 'THUMB':
        return ThumbControl(**kwargs)
    elif controlType.upper() == 'TITLEBAR':
        return TitleBarControl(**kwargs)
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


if __name__ == '__main__':
    # a = win32net.NetUserGetInfo(win32net.NetGetAnyDCName(), win32api.GetUserName(), 2)
    username = win32api.GetUserName()

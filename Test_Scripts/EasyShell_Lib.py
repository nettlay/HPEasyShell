from Library.CommonLib import QAUtils, TxtUtils, getElementByType
import os
import time

if not os.path.exists('c:\\svc'):
    os.mkdir("C:\\svc")
if not os.path.exists('c:\\svc\\svcconfig.ini'):
    os.system('echo {}\\test.txt>c:\\svc\\svcconfig.ini'.format(os.getcwd()))

file_path = TxtUtils('c:\\svc\\svcconfig.ini').get_source().strip()
file_path = os.path.dirname(file_path)
ElementlibPath = os.path.join(file_path, 'Configuration\\elementLib.ini')


def getElementMapping(filepath=ElementlibPath):
    """
    # element format:
    # define name:"name":automationid:controltype
    # eg.: OKButton:"OK":Button----->By name
    #      CancelButton:btnCancel:Button----->By automationId
    :param filepath: ElementlibPath
    :return: element
    """
    mappingDict = {}
    lines = TxtUtils(filepath).get_lines()
    for line in lines:
        if line[0] == '#':
            continue
        items = line.strip().split(":", 1)
        mappingDict[items[0]] = items[1]
    return mappingDict


def getElement(name, **kwargs):
    print(name)
    # name is defined name, format: defined name:"Name"/AutomationId:ControlType
    elementId = getElementMapping()[name].split(':')[0]
    controltype = getElementMapping()[name].split(':')[1].upper()
    if elementId.__contains__('"'):
        for i in range(3):
            if getElementByType(controltype, Name=elementId.replace('"', ''), **kwargs).Exists():
                return getElementByType(controltype, Name=elementId.replace('"', ''), **kwargs)
            else:
                time.sleep(1)
                continue
        print('{} not exist'.format(elementId))
    else:
        for i in range(3):
            if getElementByType(controltype, AutomationId=elementId, **kwargs).Exists():
                return getElementByType(controltype, AutomationId=elementId, **kwargs)
            else:
                time.sleep(1)
                continue
        print('{} not exist'.format(elementId))


class CommonUtils(QAUtils):
    @staticmethod
    def SwitchToUser():
        QAUtils.SwitchUser("User", "User", "")

    @staticmethod
    def SwitchToAdmin():
        QAUtils.SwitchUser("Admin", "Admin", "")

    @staticmethod
    def launchFromControl():
        QAUtils.LaunchAppFromControl("HP Easy Shell Configuration")

    @staticmethod
    def launchFromPath():
        QAUtils.LaunchAppFromFile("C:\\Program Files\\HP\\HP Easy Shell\\HPEasyShell.exe")

    @staticmethod
    def install(path):
        os.system('msiexec.exe /q /i {}'.format(path))


class EasyShell_Wnd:
    @staticmethod
    def MAIN_WINDOW():
        return getElement('MAIN_WINDOW')

    @staticmethod
    def TASK_SWITCHER():
        return getElement('TASK_SWITCHER')

    @staticmethod
    def WIFI_SELECTION():
        return getElement('WIFI_SELECTION')


UserKiosk_Dict = {
    'TaskSwitcher': getElement('TASK_SWITCHER'),
    # ---------Icon on the task Switcher --------------
    'WifiIcon': getElement('WifiIcon'),
    'SoundIcon': getElement('SoundIcon'),
    'HPWMIcon': getElement('HPWMIcon'),
    'SwitcherTime': getElement('SwitcherTime'),
    # -------------Tab icon--------------------------
    'UserBrowser': getElement('UserBrowser'),
    'UserTitles': getElement('UserTitles'),
    'UserSettings': getElement('UserSettings'),
    'UserAdmin': getElement('UserAdmin'),
    'UserPower': getElement('UserPower'),
    # ------- item under power button ---------------
    'Lock': getElement('Lock'),
    'Logoff': getElement('Logoff'),
    'Restart': getElement('Restart'),
    'Shutdown': getElement('Shutdown'),
    'Exit': getElement('Exit'),
    # ---------- Information at the bottom --------
    'IPAddr': getElement('IPAddr'),
    'HostName': getElement('HostName'),
    'MACAddr': getElement('MACAddr'),
    'Time': getElement('Time'),
    'CopyRight': getElement('CopyRight'),
    'Date': getElement('Date'),
    # ----------- Icon under Titles ----------------
    'UserApp': getElement('UserApp'),
    'UserConnection': getElement('UserConnection'),
    'UserStoreFront': getElement('UserStoreFront'),
    'UserWebsites': getElement('UserWebsites'),
    # ---------- for web browser ------------
    'WebHome': getElement('WebHome'),
    'UserKeyBoard': getElement('UserKeyBoard'),
    'AddressBar': getElement('AddressBar'),
    # 'WifiIcon': getElement('WifiIcon'),
    # 'SoundIcon': getElement('SoundIcon'),
    # 'HPWMIcon': getElement('HPWMIcon'),
    # -------- System Icon under Settings -------
    'SysKeyboardIcon': getElement('SysKeyboardIcon'),
    'SysDisplayIcon': getElement('SysDisplayIcon'),
    'SysMouseIcon': getElement('SysMouseIcon'),
    'SysSoundIcon': getElement('SysSoundIcon'),
    'SysRegionIcon': getElement('SysRegionIcon'),
    'SysNetworkConnIcon': getElement('SysNetworkConnIcon'),
    'SysDateTimeIcon': getElement('SysDateTimeIcon'),
    'SysEaseAccessCenterIcon': getElement('SysEaseAccessCenterIcon'),
    'SysIEIcon': getElement('SysIEIcon'),
    'SysWirelessIcon': getElement('SysWirelessIcon'),
    # -----------------------
    'WebIEClose': getElement('WebIEClose')
}

UserSettings_Dict = {
    # Button of User Settings
    "KioskMode": getElement('KioskMode'),
    "AllowUserSetting": getElement('AllowUserSetting'),
    'AllowMouse': getElement('AllowMouse'),
    'AllowKeyboard': getElement('AllowKeyboard'),
    'AllowDisplay': getElement('AllowDisplay'),
    'AllowSound': getElement('AllowSound'),
    'AllowRegion': getElement('AllowRegion'),
    'AllowNetworkConn': getElement('AllowNetworkConn'),
    'AllowDateTime': getElement('AllowDateTime'),
    'AllowEasyAccess': getElement('AllowEasyAccess'),
    'AllowIEProperty': getElement('AllowIEProperty'),
    'AllowWifiConfig': getElement('AllowWifiConfig'),

}

UserInterface_Dict = {
    """
    Admin settings
    """
    "KioskMode": getElement('KioskMode'),
    "DisplayTitle": getElement('DisplayTitle'),
    "DisplayApp": getElement('DisplayApp'),
    "DisplayConnections": getElement('DisplayConnections'),
    "DisplayStoreFront": getElement('DisplayStoreFront'),
    "DisplayWebsites": getElement('DisplayWebsites'),
    'DisplayBrowser': getElement('DisplayBrowser'),
    'DisplayAddress': getElement('DisplayAddress'),
    'DisplayNavigation': getElement('DisplayNavigation'),
    'DisplayHome': getElement('DisplayHome'),
    'DisplayAdmin': getElement('DisplayAdmin'),
    'DisplayPower': getElement('DisplayPower'),
    'AllowLock': getElement('AllowLock'),
    'AllowLogoff': getElement('AllowLogoff'),
    'AllowRestart': getElement('AllowRestart'),
    'AllowShutDown': getElement('AllowShutDown'),
    'DisplayVKeyboard': getElement('DisplayVKeyboard'),
    'EnableLTKeyboard': getElement('EnableLTKeyboard'),
    'DisplayTime': getElement('DisplayTime'),
    'DisplayIP': getElement('DisplayIP'),
    'DisplayMAC': getElement('DisplayMAC'),
    'EnableTaskSwitcher': getElement('EnableTaskSwitcher'),
    'Permanent': getElement('Permanent'),
    'DisplaySwitcherTime': getElement('DisplaySwitcherTime'),
    'DisplayBattery': getElement('DisplayBattery'),
    'DisplayCellular': getElement('DisplayCellular'),
    'DisplaySound': getElement('DisplaySound'),
    'DisplaySoundIconInteraction': getElement('DisplaySoundIconInteraction'),
    'DisplayWifi': getElement('DisplayWifi'),
    'DisplayWifiInterAction': getElement('DisplayWifiInterAction'),
    'DisplayWriteFilter': getElement('DisplayWriteFilter'),
    'DisplayWriteFilterInteraction': getElement('DisplayWriteFilterInteraction'),
    'HideEasyShell': getElement('HideEasyShell'),
    'EnableCustom': getElement('EnableCustom'),
    'DisplayNetwork': getElement('DisplayNetwork'),
    'EnableSmartcard': getElement('EnableSmartcard'),
    # user settings
    "AllowUserSetting": getElement('AllowUserSetting'),
    'AllowMouse': getElement('AllowMouse'),
    'AllowKeyboard': getElement('AllowKeyboard'),
    'AllowDisplay': getElement('AllowDisplay'),
    'AllowSound': getElement('AllowSound'),
    'AllowRegion': getElement('AllowRegion'),
    'AllowNetworkConn': getElement('AllowNetworkConn'),
    'AllowDateTime': getElement('AllowDateTime'),
    'AllowEasyAccess': getElement('AllowEasyAccess'),
    'AllowIEProperty': getElement('AllowIEProperty'),
    'AllowWifiConfig': getElement('AllowWifiConfig'),
}

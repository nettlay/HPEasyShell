from QAUIutils import QATools
import QAUIutils as QAUtil
import uiautomation
import winreg


class Common(QATools):
    def __init__(self):
        QATools.__init__(self)

    MAIN_WINDOW = QAUtil.WindowMethod(Name="HP Easy Shell")
    Task_Switcher = QAUtil.WindowMethod(Name='HP Easy Shell Application Switcher')
    WifiSelection = QAUtil.WindowMethod(Name='HP Wireless Configuration')

    @staticmethod
    def AddProxy(user):
        """
         User must be actual user that exist in windows
        """
        sidDict = QATools.getSid()
        sid = sidDict[user]
        root = winreg.HKEY_USERS
        path = r"{}\Software\Microsoft\Windows\CurrentVersion\Internet Settings".format(sid)
        key = winreg.OpenKeyEx(root, path, access=winreg.KEY_ALL_ACCESS | winreg.KEY_WOW64_64KEY)
        winreg.SetValueEx(key, 'ProxyServer', 0, winreg.REG_SZ, '15.85.199.199:8080')
        winreg.SetValueEx(key, 'ProxyOverride', 0, winreg.REG_SZ, '*.sh.dto;15.83.*.*')
        winreg.SetValueEx(key, 'ProxyEnable', 0, winreg.REG_DWORD, 1)

    @staticmethod
    def launchFromControl():
        QATools.LaunchAppFromControl("HP Easy Shell Configuration")

    @staticmethod
    def launchFromPath():
        QATools.LaunchAppFromFile("C:\\Program Files\\HP\\HP Easy Shell\\HPEasyShell.exe")

    @staticmethod
    def SwitchToUser():
        QATools.SwitchUser("User", "User", "")

    @staticmethod
    def SwitchToAdmin():
        QATools.SwitchUser("Admin", "Admin", "")

    @staticmethod
    def GetFileList(file):
        return QATools.GetListFromFile(file)


class ButtonList:
    """
    Define all the Button element
    """

    def __init__(self):
        pass

    @staticmethod
    def Common(**kwargs):
        return QAUtil.ButtonMethod(**kwargs)

    @staticmethod
    def UserBrowser(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='UserSitesButton', Name='', **kwargs)

    @staticmethod
    def UserAdmin(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='UserAdminButton', Name='', **kwargs)

    @staticmethod
    def UserSettings(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='UserSettingsButton', **kwargs)

    @staticmethod
    def UserPower(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='UserpowerButton', Name='', **kwargs)

    # /////    button shown on switcher window
    @staticmethod
    def WifiIcon(**kwargs):
        return QAUtil.ButtonMethod(searchFromControl=Common.Task_Switcher, AutomationId='WifiButton', **kwargs)

    @staticmethod
    def SoundIcon(**kwargs):
        return QAUtil.ButtonMethod(searchFromControl=Common.Task_Switcher, AutomationId='SoundButton', **kwargs)

    @staticmethod
    def SoundAdjust(**kwargs):
        return QAUtil.Method(searchFromControl=Common.Task_Switcher, AutomationId='SoundSlider', **kwargs)

    @staticmethod
    def SoundAdjustBar(**kwargs):
        return QAUtil.Method(searchFromControl=Common.Task_Switcher, AutomationId='Thumb', **kwargs)

    @staticmethod
    def HPWMIcon(**kwargs):
        return QAUtil.ButtonMethod(searchFromControl=Common.Task_Switcher, AutomationId='WriteFilterButton', **kwargs)

    @staticmethod
    def UserKeyBoard(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='keyboardButton', **kwargs)

    @staticmethod
    def AdminPower(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='powerButton', **kwargs)

    @staticmethod
    def Settings(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='settingsButton', **kwargs)

    @staticmethod
    def Applications(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='applicationsButton', **kwargs)

    @staticmethod
    def Connections(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='connectionsButton', **kwargs)

    @staticmethod
    def StoreFront(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='CitrixButton', Name='', **kwargs)

    @staticmethod
    def WebSites(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='websitesButton', Name='', **kwargs)

    @staticmethod
    def KioskMode(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='KioskModeToggle123', Name='', **kwargs)

    @staticmethod
    def APPLY(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='ApplyButton', Name='', **kwargs)

    @staticmethod
    def WebHome(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='WebHomeButton', Name='', **kwargs)

    @staticmethod
    def DisplayTitle(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='DisplayTilesToggle', Name='', **kwargs)

    @staticmethod
    def DisplayApp(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='DisplayApplicationsToggle', Name='', **kwargs)

    @staticmethod
    def DisplayHome(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='DisplayHomeToggle', Name='', **kwargs)

    @staticmethod
    def DisplayNavigation(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='DisplayNavToggle', Name='', **kwargs)

    @staticmethod
    def DisplayAddress(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='DisplayAddressToggle', Name='', **kwargs)

    @staticmethod
    def DisplayBrowser(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='DisplayBrowserToggle', Name='', **kwargs)

    @staticmethod
    def DisplayWebsites(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='DisplayWebsitesToggle', Name='', **kwargs)

    @staticmethod
    def DisplayStoreFront(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='DisplayCitrixToggle', Name='', **kwargs)

    @staticmethod
    def DisplayConnections(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='DisplayConnectionsToggle', Name='', **kwargs)

    @staticmethod
    def DisplayAdmin(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='DisplayAdminToggle', Name='', **kwargs)

    @staticmethod
    def DisplayPower(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='DisplayPowerToggle', Name='', **kwargs)

    @staticmethod
    def AllowLock(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='AllowLockToggle', Name='', **kwargs)

    @staticmethod
    def AllowLogoff(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='AllowLogoffToggle', Name='', **kwargs)

    @staticmethod
    def AllowRestart(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='AllowRebootToggle', Name='', **kwargs)

    @staticmethod
    def AllowShutDown(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='AllowShutdownToggle', Name='', **kwargs)

    @staticmethod
    def DisplayVKeyboard(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='DisplayKeyboardToggle', Name='', **kwargs)

    @staticmethod
    def EnableLTKeyboard(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='virtualKeyboardLegacySupportToggle', Name='', **kwargs)

    @staticmethod
    def DisplayTime(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='DisplayTimeToggle', Name='', **kwargs)

    @staticmethod
    def DisplayIP(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='DisplayIPToggle', Name='', **kwargs)

    @staticmethod
    def DisplayMAC(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='DisplayMacAddressToggle', Name='', **kwargs)

    @staticmethod
    def EnableTaskSwitcher(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='DisplayThinBarToggle', Name='', **kwargs)

    @staticmethod
    def Permanent(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='taskSwicherPermanentToggle', **kwargs)

    @staticmethod
    def DisplaySwitcherTime(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='DisplayTimeButtonToggle', **kwargs)

    @staticmethod
    def DisplayBattery(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='DisplayBatteryToggle', **kwargs)

    @staticmethod
    def DisplayCellular(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='DisplayMobileToggle', **kwargs)

    @staticmethod
    def DisplaySound(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='DisplaySoundToggle', **kwargs)

    @staticmethod
    def DisplaySoundIconInteraction(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='SoundInteractionToggle', **kwargs)

    @staticmethod
    def DisplayWifi(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='DisplayWifiToggle', **kwargs)

    @staticmethod
    def DisplayWifiInterAction(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='WifiInteractionToggle', **kwargs)

    @staticmethod
    def DisplayWriteFilterInteraction(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='WriteFilterInteractionToggle', **kwargs)

    @staticmethod
    def DisplayWriteFilter(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='DisplayWriteFilterToggle', **kwargs)

    @staticmethod
    def DisplayNetwork(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='EnableNetworkStatusToggle', **kwargs)

    @staticmethod
    def HideEasyShell(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='HideThinShellToggle', **kwargs)

    @staticmethod
    def EnableCustom(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='CustomBackground', **kwargs)

    @staticmethod
    def BGFileLocation(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='bgFileLocationButton', **kwargs)

    @staticmethod
    def SelectTheme(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='ThemeButton', **kwargs)

    @staticmethod
    def AllowUserSetting(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='AllowUserSettingsToggle', **kwargs)

    @staticmethod
    def AllowMouse(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='MouseToggle', **kwargs)

    @staticmethod
    def AllowKeyboard(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='KeyboardToggle', **kwargs)

    @staticmethod
    def AllowDisplay(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='DisplayToggle', **kwargs)

    @staticmethod
    def AllowSound(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='SoundToggle', **kwargs)

    @staticmethod
    def AllowRegion(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='RegionLanguageToggle', **kwargs)

    @staticmethod
    def AllowNetworkConn(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='NetworkConnectionsToggle', **kwargs)

    @staticmethod
    def AllowDateTime(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='DateTimeToggle', **kwargs)

    @staticmethod
    def AllowEasyAccess(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='EaseAccessCenterToggle', **kwargs)

    @staticmethod
    def AllowIEProperty(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='IEPropertiesToggle', **kwargs)

    @staticmethod
    def AllowWifiConfig(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='WifiConfigToggle', **kwargs)

    @staticmethod
    def Advanced(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='advancedSettingsButton', **kwargs)

    @staticmethod
    def EnableSmartcard(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='EnableSmartCardActionToggle', **kwargs)

    @staticmethod
    def RebootScheduler(**kwargs):
        return QAUtil.ButtonMethod(Name='Reboot Scheduler', **kwargs)

    @staticmethod
    def Export(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='buttonExport', **kwargs)

    @staticmethod
    def Exit(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='ExitButton', **kwargs)

    @staticmethod
    def Default(**kwargs):
        # /////////////// This is a TextControl, It is used to click its parent button ///////////////////
        return QAUtil.TextMethod(AutomationId='CurrentProfile', **kwargs)

    # /////////// Buttons of Kiosk user mode /////////////////////////
    @staticmethod
    def UserTitles(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='UserDefaultButton', **kwargs)

    # /////////////Buttons of app edit under applications//////////////////////////

    @staticmethod
    def ApplicationAdd(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='appButton', **kwargs)

    @staticmethod
    def AppPathBrowser(**kwargs):
        return QAUtil.ButtonMethod(Name='...', foundIndex=1, **kwargs)

    @staticmethod
    def AppCustomIconBrowser(**kwargs):
        return QAUtil.ButtonMethod(Name='...', foundIndex=2, **kwargs)

    @staticmethod
    def AppCustomIconSwitch(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='customIconToggle', **kwargs)

    @staticmethod
    def AppAutoLaunch(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='ToggleAutoLaunch', **kwargs)

    @staticmethod
    def AppPersistent(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='TogglePersistent', **kwargs)

    @staticmethod
    def AppMaximized(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='ToggleMaximized', **kwargs)

    @staticmethod
    def AppAdminOnly(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='ToggleAdminOnly', **kwargs)

    @staticmethod
    def AppHideMissing(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='HideTileToggle', **kwargs)

    @staticmethod
    def AppOK(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='SaveButton', **kwargs)

    @staticmethod
    def APPCancel(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='CancelButton', **kwargs)

    @staticmethod
    def DeleteYes(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='6', **kwargs)

    @staticmethod
    def DeleteNo(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='7', **kwargs)

    # //////////////////button of RDP under connection////////////
    @staticmethod
    def RDPAdd(**kwargs):
        """
        Find this button by its name, then return parentControl
        """
        txt = QAUtil.TextMethod(Name='RDP', **kwargs)
        return txt.GetParentControl()

    @staticmethod
    def VMwareAdd(**kwargs):
        """
        Find this button by its name, then return parentControl
        """
        txt = QAUtil.TextMethod(Name='VMware', **kwargs)
        return txt.GetParentControl()

    @staticmethod
    def CitrixICAAdd(**kwargs):
        """
        Find this button by its name, then return parentControl
        """
        txt = QAUtil.TextMethod(Name='CitrixICA', **kwargs)
        return txt.GetParentControl()

    @staticmethod
    def WebsiteAdd(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='protocolButton', **kwargs)

    @staticmethod
    def StoreFrontAdd(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='storefrontButton', **kwargs)

    @staticmethod
    def WebIEClose(**kwargs):
        return QAUtil.ButtonMethod(AutomationId='WebIECloseButton', **kwargs)

class MenuItemList:
    @staticmethod
    def Common(**kwargs):
        return QAUtil.Method(ControlType=uiautomation.ControlType.MenuItemControl,
                             **kwargs)

    @staticmethod
    def Lock(**kwargs):
        return QAUtil.Method(ControlType=uiautomation.ControlType.MenuItemControl,
                             Name='Lock', **kwargs)

    @staticmethod
    def Logoff(**kwargs):
        return QAUtil.Method(ControlType=uiautomation.ControlType.MenuItemControl, Name='Log off', **kwargs)

    @staticmethod
    def Restart(**kwargs):
        return QAUtil.Method(ControlType=uiautomation.ControlType.MenuItemControl, Name='Restart', **kwargs)

    @staticmethod
    def Shutdown(**kwargs):
        return QAUtil.Method(ControlType=uiautomation.ControlType.MenuItemControl, Name='Shut down', **kwargs)

    @staticmethod
    def Exit(**kwargs):
        return QAUtil.Method(ControlType=uiautomation.ControlType.MenuItemControl, Name='Exit', **kwargs)


class TextList:
    @staticmethod
    def Common(**kwargs):
        return QAUtil.TextMethod(**kwargs)

    @staticmethod
    def CopyRight(**kwargs):
        return QAUtil.TextMethod(AutomationId='copyrightTextBlock', **kwargs)

    @staticmethod
    def Time(**kwargs):
        return QAUtil.TextMethod(searchFromControl=Common.MAIN_WINDOW, AutomationId='LabelTime', **kwargs)

    @staticmethod
    def Date(**kwargs):
        return QAUtil.TextMethod(AutomationId='LabelDate', **kwargs)

    @staticmethod
    def IPAddr(**kwargs):
        return QAUtil.TextMethod(AutomationId='LabelIP', **kwargs)

    @staticmethod
    def MACAddr(**kwargs):
        return QAUtil.TextMethod(AutomationId='LabelMacAddress', **kwargs)

    @staticmethod
    def HostName(**kwargs):
        return QAUtil.TextMethod(AutomationId='LabelMachineName', **kwargs)

    @staticmethod
    def UserApp(**kwargs):
        return QAUtil.TextMethod(AutomationId='xLabelAppTitle', **kwargs)

    @staticmethod
    def UserConnection(**kwargs):
        return QAUtil.TextMethod(AutomationId='xLabelConnectionTitle', **kwargs)

    @staticmethod
    def UserStoreFront(**kwargs):
        return QAUtil.TextMethod(AutomationId='xLabelCitrixTitle', **kwargs)

    @staticmethod
    def UserWebsites(**kwargs):
        return QAUtil.TextMethod(AutomationId='xLabelWebTitle', **kwargs)

    # ---------------Icon on task swither -------------------------------------
    @staticmethod
    def SwitcherTime(**kwargs):
        return QAUtil.TextMethod(searchFromControl=Common.Task_Switcher, AutomationId='LabelTime', **kwargs)

    # -------------- System Icon under User Settings ------------------------
    @staticmethod
    def SysKeyboardIcon(**kwargs):
        return QAUtil.TextMethod(Name='Keyboard', **kwargs)

    @staticmethod
    def SysDisplayIcon(**kwargs):
        return QAUtil.TextMethod(Name='Display', **kwargs)

    @staticmethod
    def SysMouseIcon(**kwargs):
        return QAUtil.TextMethod(Name='Mouse', **kwargs)

    @staticmethod
    def SysSoundIcon(**kwargs):
        return QAUtil.TextMethod(Name='Sound', **kwargs)

    @staticmethod
    def SysRegionIcon(**kwargs):
        return QAUtil.TextMethod(Name='Region and Language', **kwargs)

    @staticmethod
    def SysNetworkConnIcon(**kwargs):
        return QAUtil.TextMethod(Name='Network Connections', **kwargs)

    @staticmethod
    def SysDateTimeIcon(**kwargs):
        return QAUtil.TextMethod(Name='Date and Time', **kwargs)

    @staticmethod
    def SysEaseAccessCenterIcon(**kwargs):
        return QAUtil.TextMethod(Name='Ease of Access Center', **kwargs)

    @staticmethod
    def SysIEIcon(**kwargs):
        return QAUtil.TextMethod(Name='Internet Properties', **kwargs)

    @staticmethod
    def SysWirelessIcon(**kwargs):
        return QAUtil.TextMethod(Name='Wireless configuration', **kwargs)


class EditList:
    @staticmethod
    def AddressBar(**kwargs):
        return QAUtil.EditMethod(AutomationId='addressBar', **kwargs)

    @staticmethod
    def SaveToFile(**kwargs):
        return QAUtil.EditMethod(AutomationId="1001", **kwargs)

    @staticmethod
    def BGFileLocation(**kwargs):
        return QAUtil.EditMethod(AutomationId="bgFileLocation", **kwargs)

    # /////App edit windows element/////////////////
    @staticmethod
    def AppName(**kwargs):
        return QAUtil.EditMethod(AutomationId='TextBoxAppName', **kwargs)

    @staticmethod
    def AppPath(**kwargs):
        return QAUtil.EditMethod(AutomationId='TextBoxAppPath', **kwargs)

    @staticmethod
    def AppArguments(**kwargs):
        return QAUtil.EditMethod(AutomationId='TextBoxAppArguments', **kwargs)

    @staticmethod
    def AppLaunchDelay(**kwargs):
        return QAUtil.EditMethod(AutomationId='TextBoxAppLaunchDelay', **kwargs)

    @staticmethod
    def AppCustomIcon(**kwargs):
        return QAUtil.EditMethod(AutomationId='iconPathTextBox', **kwargs)


class DataGridList:
    @staticmethod
    def ProfileDG(**kwargs):
        return QAUtil.DataGridMethod(AutomationId="lvProfiles", **kwargs)


class TabItemList:
    # /////////////app edit under Applications ///////////
    @staticmethod
    def AppBasic(**kwargs):
        return QAUtil.Method(ControlType=uiautomation.ControlType.TabItemControl, Name='Basic', **kwargs)

    @staticmethod
    def AppBehaviors(**kwargs):
        return QAUtil.Method(ControlType=uiautomation.ControlType.TabItemControl, Name='Behaviors', **kwargs)

    # ///////////// RDP Edit under connections /////////
    @staticmethod
    def RDPLocalResources(**kwargs):
        return QAUtil.Method(ControlType=uiautomation.ControlType.TabItemControl, Name='Local resources', **kwargs)

    @staticmethod
    def RDPPrograms(**kwargs):
        return QAUtil.Method(ControlType=uiautomation.ControlType.TabItemControl, Name='Programs', **kwargs)

    @staticmethod
    def RDPExperience(**kwargs):
        return QAUtil.Method(ControlType=uiautomation.ControlType.TabItemControl, Name='Experience', **kwargs)

    @staticmethod
    def RDPAdvanced(**kwargs):
        return QAUtil.Method(ControlType=uiautomation.ControlType.TabItemControl, Name='Advanced', **kwargs)

    @staticmethod
    def RDPExpert(**kwargs):
        return QAUtil.Method(ControlType=uiautomation.ControlType.TabItemControl, Name='Expert', **kwargs)

    @staticmethod
    def RDPBehaviors(**kwargs):
        return QAUtil.Method(ControlType=uiautomation.ControlType.TabItemControl, Name='Behaviors', **kwargs)

    # ////////////////VMware edit under connections ///////////////
    @staticmethod
    def VMBasic(**kwargs):
        return QAUtil.Method(ControlType=uiautomation.ControlType.TabItemControl, Name='Basic', **kwargs)

    @staticmethod
    def VMAdvanced(**kwargs):
        return QAUtil.Method(ControlType=uiautomation.ControlType.TabItemControl, Name='Advanced', **kwargs)

    @staticmethod
    def VMBehaviors(**kwargs):
        return QAUtil.Method(ControlType=uiautomation.ControlType.TabItemControl, Name='Behaviors', **kwargs)

    # /////////////////Citrix edit under connections //////////////////////
    @staticmethod
    def CitrixBasic(**kwargs):
        return QAUtil.Method(ControlType=uiautomation.ControlType.TabItemControl, Name='Basic', **kwargs)

    @staticmethod
    def CitrixDisplay(**kwargs):
        return QAUtil.Method(ControlType=uiautomation.ControlType.TabItemControl, Name='Display', **kwargs)

    @staticmethod
    def CitrixConnection(**kwargs):
        return QAUtil.Method(ControlType=uiautomation.ControlType.TabItemControl, Name='Connection options', **kwargs)

    @staticmethod
    def CitrixAdvanced(**kwargs):
        return QAUtil.Method(ControlType=uiautomation.ControlType.TabItemControl, Name='Advanced', **kwargs)

    @staticmethod
    def CitrixBehaviors(**kwargs):
        return QAUtil.Method(ControlType=uiautomation.ControlType.TabItemControl, Name='Behaviors', **kwargs)

    # //////////////////Store front Edit under connections //////////////////////////
    @staticmethod
    def StoreFront(**kwargs):
        return QAUtil.Method(ControlType=uiautomation.ControlType.TabItemControl, Name='StoreFront', **kwargs)

    @staticmethod
    def StoreOptions(**kwargs):
        return QAUtil.Method(ControlType=uiautomation.ControlType.TabItemControl, Name='Options', **kwargs)

    @staticmethod
    def StoreBehaviors(**kwargs):
        return QAUtil.Method(ControlType=uiautomation.ControlType.TabItemControl, Name='Behaviors', **kwargs)

    # /////////////////Website edit under websites /////////////////
    @staticmethod
    def WebBasic(**kwargs):
        return QAUtil.Method(ControlType=uiautomation.ControlType.TabItemControl, Name='Basic', **kwargs)

    @staticmethod
    def WebBehaviors(**kwargs):
        return QAUtil.Method(ControlType=uiautomation.ControlType.TabItemControl, Name='Behaviors', **kwargs)


UserKiosk_Dict = {
    'TaskSwitcher': Common.Task_Switcher,
    # ---------Icon on the task Switcher --------------
    'WifiIcon': ButtonList.WifiIcon(),
    'SoundIcon': ButtonList.SoundIcon(),
    'HPWMIcon': ButtonList.HPWMIcon(),
    'SwitcherTime': TextList.SwitcherTime(),
    # ---------------------------------------
    'UserBrowser': ButtonList.UserBrowser(),
    'UserTitles': ButtonList.UserTitles(),
    'UserSettings': ButtonList.UserSettings(),
    'UserAdmin': ButtonList.UserAdmin(),
    'UserPower': ButtonList.UserPower(),
    # ------- item under power button ---------------
    'Lock': MenuItemList.Lock(),
    'Logoff': MenuItemList.Logoff(),
    'Restart': MenuItemList.Restart(),
    'Shutdown': MenuItemList.Shutdown(),
    'Exit': MenuItemList.Exit(),
    # ---------- Information at the bottom --------
    'IPAddr': TextList.IPAddr(),
    'HostName': TextList.HostName(),
    'MACAddr': TextList.MACAddr(),
    'Time': TextList.Time(),
    'CopyRight': TextList.CopyRight(),
    'Date': TextList.Date(),
    # ----------- Icon under Titles ----------------
    'UserApp': TextList.UserApp(),
    'UserConnection': TextList.UserConnection(),
    'UserStoreFront': TextList.UserStoreFront(),
    'UserWebsites': TextList.UserWebsites(),
    # ---------- for web browser ------------
    'WebHome': ButtonList.WebHome(),
    'UserKeyBoard': ButtonList.UserKeyBoard(),
    'AddressBar': EditList.AddressBar(),
    'WifiIcon': ButtonList.WifiIcon(),
    'SoundIcon': ButtonList.SoundIcon(),
    'HPWMIcon': ButtonList.HPWMIcon(),
    # -------- System Icon under Settings -------
    'SysKeyboardIcon': TextList.SysKeyboardIcon(),
    'SysDisplayIcon': TextList.SysDisplayIcon(),
    'SysMouseIcon': TextList.SysMouseIcon(),
    'SysSoundIcon': TextList.SysSoundIcon(),
    'SysRegionIcon': TextList.SysRegionIcon(),
    'SysNetworkConnIcon': TextList.SysNetworkConnIcon(),
    'SysDateTimeIcon': TextList.SysDateTimeIcon(),
    'SysEaseAccessCenterIcon': TextList.SysEaseAccessCenterIcon(),
    'SysIEIcon': TextList.SysIEIcon(),
    'SysWirelessIcon': TextList.SysWirelessIcon(),
    # -----------------------
    'WebIEClose': ButtonList.WebIEClose()
}

UserSettings_Dict = {
    # Button of User Settings
    "AllowUserSetting": ButtonList.AllowUserSetting(),
    'AllowMouse': ButtonList.AllowMouse(),
    'AllowKeyboard': ButtonList.AllowKeyboard(),
    'AllowDisplay': ButtonList.AllowDisplay(),
    'AllowSound': ButtonList.AllowSound(),
    'AllowRegion': ButtonList.AllowRegion(),
    'AllowNetworkConn': ButtonList.AllowNetworkConn(),
    'AllowDateTime': ButtonList.AllowDateTime(),
    'AllowEasyAccess': ButtonList.AllowEasyAccess(),
    'AllowIEProperty': ButtonList.AllowIEProperty(),
    'AllowWifiConfig': ButtonList.AllowWifiConfig(),
    # Button of User Interface
    "DisplayTitle": ButtonList.DisplayTitle(),
    "DisplayApp": ButtonList.DisplayApp(),
    "DisplayConnections": ButtonList.DisplayConnections(),
    "DisplayStoreFront": ButtonList.DisplayStoreFront(),
    "DisplayWebsites": ButtonList.DisplayWebsites(),
    'DisplayBrowser': ButtonList.DisplayBrowser(),
    'DisplayAddress': ButtonList.DisplayAddress(),
    'DisplayNavigation': ButtonList.DisplayNavigation(),
    'DisplayHome': ButtonList.DisplayHome(),
    'DisplayAdmin': ButtonList.DisplayAdmin(),
    'DisplayPower': ButtonList.DisplayPower(),
    'AllowLock': ButtonList.AllowLock(),
    'AllowLogoff': ButtonList.AllowLogoff(),
    'AllowRestart': ButtonList.AllowRestart(),
    'AllowShutDown': ButtonList.AllowShutDown(),
    'DisplayVKeyboard': ButtonList.DisplayVKeyboard(),
    'EnableLTKeyboard': ButtonList.EnableLTKeyboard(),
    'DisplayTime': ButtonList.DisplayTime(),
    'DisplayIP': ButtonList.DisplayIP(),
    'DisplayMAC': ButtonList.DisplayMAC(),
    'EnableTaskSwitcher': ButtonList.EnableTaskSwitcher(),
    'Permanent': ButtonList.Permanent(),
    'DisplaySwitcherTime': ButtonList.DisplaySwitcherTime(),
    'DisplayBattery': ButtonList.DisplayBattery(),
    'DisplayCellular': ButtonList.DisplayCellular(),
    'DisplaySound': ButtonList.DisplaySound(),
    'DisplaySoundIconInteraction': ButtonList.DisplaySoundIconInteraction(),
    'DisplayWifi': ButtonList.DisplayWifi(),
    'DisplayWifiInterAction': ButtonList.DisplayWifiInterAction(),
    'DisplayWriteFilter': ButtonList.DisplayWriteFilter(),
    'DisplayWriteFilterInteraction': ButtonList.DisplayWriteFilterInteraction(),
    'HideEasyShell': ButtonList.HideEasyShell(),
    'EnableCustom': ButtonList.EnableCustom(),
    'DisplayNetwork': ButtonList.DisplayNetwork(),
    'EnableSmartcard': ButtonList.EnableSmartcard(),
}
UserInterface_Dict = {
    """
    Admin settings
    """
    "DisplayTitle": ButtonList.DisplayTitle(),
    "DisplayApp": ButtonList.DisplayApp(),
    "DisplayConnections": ButtonList.DisplayConnections(),
    "DisplayStoreFront": ButtonList.DisplayStoreFront(),
    "DisplayWebsites": ButtonList.DisplayWebsites(),
    'DisplayBrowser': ButtonList.DisplayBrowser(),
    'DisplayAddress': ButtonList.DisplayAddress(),
    'DisplayNavigation': ButtonList.DisplayNavigation(),
    'DisplayHome': ButtonList.DisplayHome(),
    'DisplayAdmin': ButtonList.DisplayAdmin(),
    'DisplayPower': ButtonList.DisplayPower(),
    'AllowLock': ButtonList.AllowLock(),
    'AllowLogoff': ButtonList.AllowLogoff(),
    'AllowRestart': ButtonList.AllowRestart(),
    'AllowShutDown': ButtonList.AllowShutDown(),
    'DisplayVKeyboard': ButtonList.DisplayVKeyboard(),
    'EnableLTKeyboard': ButtonList.EnableLTKeyboard(),
    'DisplayTime': ButtonList.DisplayTime(),
    'DisplayIP': ButtonList.DisplayIP(),
    'DisplayMAC': ButtonList.DisplayMAC(),
    'EnableTaskSwitcher': ButtonList.EnableTaskSwitcher(),
    'Permanent': ButtonList.Permanent(),
    'DisplaySwitcherTime': ButtonList.DisplaySwitcherTime(),
    'DisplayBattery': ButtonList.DisplayBattery(),
    'DisplayCellular': ButtonList.DisplayCellular(),
    'DisplaySound': ButtonList.DisplaySound(),
    'DisplaySoundIconInteraction': ButtonList.DisplaySoundIconInteraction(),
    'DisplayWifi': ButtonList.DisplayWifi(),
    'DisplayWriteFilter': ButtonList.DisplayWriteFilter(),
    'HideEasyShell': ButtonList.HideEasyShell(),
    'EnableCustom': ButtonList.EnableCustom(),
}

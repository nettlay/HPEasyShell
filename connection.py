import Test_Scripts.EasyShell_Lib as EasyshellLib
import Library.CommonLib as CommonLib
import os
import time
import traceback

"""
interaction command between local and remote session:
local file: test_vdi.txt (ex. test_rdp.txt)
remote file: test_vdi_result.txt (ex. test_rdp_result.txt)
test logon: send by local, remote get this command should logoff session
test xxxxx: send by local, remote get this command should test function xxxxx
test status: send by local, remote get this command should return vdi's test status 
"test_status_Finished | test_status_testing | None"
"""


class Logon:
    def __init__(self):
        self.session_pane = 'RDPSessionPane'
        self.local_file = 'test_citrix.txt'
        self.remote_file = 'test_citrix_result.txt'
        self.log_path = '/Function/Automation/log/citrix'
        self.ftp_token = {'ip': '15.83.251.201', 'username': 'sh\\cheng.balance', 'password': 'password.321'}

    def input_credential(self):
        pass

    @staticmethod
    def install_ca(file):
        EasyshellLib.CommonUtils.import_cert(file)

    @staticmethod
    def check_cert(name):
        return EasyshellLib.CommonUtils.get_cert_id(name)

    @staticmethod
    def del_cert(index):
        EasyshellLib.CommonUtils.del_cert(index)

    def get_resolution(self):
        rect = EasyshellLib.getElement(self.session_pane).BoundingRectangle
        return rect[2] - rect[0], rect[3] - rect[1]

    def utils(self, profile='', op='exist', item='normal'):
        """
        :param profile:  test profile, [test1,test2,,standardApp...]
        :param op: test option [exist | notexist | shown | edit | delete |launch |default]
        :param item: specific for connection, if item=connection, connection button element=textcontrol.getparent,
                else element=textcontrol.getparent.getparent
        :return: Bool
        """
        name = profile["Name"]
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

    def get_test_status(self, ftp, wait_time):
        flag = None
        ftp.change_dir(self.log_path)
        for i in range(wait_time):
            if self.remote_file in ftp.get_item_list('.'):
                try:
                    ftp.download_file(self.remote_file, self.remote_file)
                except:
                    print('Download remote file fail')
                    time.sleep(1)
                    continue
                with open(self.remote_file, 'r') as f:
                    status = f.read().split('_')[-1]
                os.remove(self.remote_file)
                flag = status
                break
            else:
                flag = 'FAIL'
                time.sleep(1)
                continue
        try:
            ftp.delete_file(self.remote_file)
            # ftp.delete_file(self.local_file)
        except:
            print('delete ftp file fail')
        try:
            flag = flag.strip()
            return flag
        except:
            return flag

    @staticmethod
    def wait_element(element, cycles=5, exists=True):
        flag = None
        if exists:
            for i in range(cycles):
                if element.Exists(0, 0):
                    flag = element
                    break
                else:
                    flag = None
                    print('Exist loop %d' % i)
                    time.sleep(1)
                    continue
        else:
            for i in range(cycles):
                if element.Exists(0, 0):
                    print('Not exist loop %d' % i)
                    time.sleep(1)
                    continue
                else:
                    flag = True
                    break
        return flag

    def logon(self, profile):
        pass

    def wait_logon(self):
        pass

    def check_error(self, wnd):
        return

    def check_logon(self, wait_time):
        ftp = CommonLib.FTPUtils(self.ftp_token['ip'], self.ftp_token['username'], self.ftp_token['password'])
        for i in range(wait_time):
            ftp.upload_file(self.local_file, self.local_file)
        ftp.close()

    def get_feedback(self):
        ftp = CommonLib.FTPUtils(self.ftp_token['ip'], self.ftp_token['username'], self.ftp_token['password'])
        if self.remote_file in ftp.get_item_list(self.log_path):
            ftp.download_file(self.remote_file, self.remote_file)
            ftp.close()
            return True
        else:
            ftp.close()
            return False

    def logoff(self):
        os.system('echo test_logon_pass>{}'.format(self.remote_file))
        ftp = CommonLib.FTPUtils(self.ftp_token['ip'], self.ftp_token['username'], self.ftp_token['password'])
        ftp.change_dir(self.log_path)
        ftp.upload_file(self.remote_file, self.remote_file)
        os.remove(self.remote_file)
        ftp.close()
        os.system('shutdown -l')


class RDPLogon(Logon):
    def __init__(self):
        Logon.__init__(self)
        self.session_pane = 'RDPSessionPane'
        self.local_file = 'test_rdp.txt'
        self.remote_file = 'test_rdp_result.txt'
        self.log_path = '/Function/Automation/log/rdp'

    def check_resolution(self):
        pass

    def logon(self, profile):
        if profile['Autolaunch'] == 'OFF' or not profile['Autolaunch']:
            self.utils(profile, 'launch', 'conn')
            # check connection not launch directly when launchdelay 30
        time.sleep(profile['Launchdelay'])
        time.sleep(2)
        if self.wait_element(
                EasyshellLib.getElement('Connect', searchFromControl=EasyshellLib.getElement('RDP_WARNNING'))):
            EasyshellLib.getElement('Connect',
                                    searchFromControl=EasyshellLib.getElement('RDP_WARNNING')).Click()
        else:
            print('-----------Manual Launch fail, connection not launch------------------------')
            return False
        # loop 3 times to wait for warning dialog pops up
        time.sleep(2)
        print('等待密码输入框, 密码: ' + profile['Password'])
        if self.wait_element(EasyshellLib.getElement('RDPPassword')):
            EasyshellLib.getElement('RDPPassword').SetValue(profile['Password'])
            EasyshellLib.getElement('ButtonOK').Click()
        # loop 3 times wait for certificate dialog pops up
        time.sleep(2)
        print('等待证书提示框')
        if self.wait_element(EasyshellLib.getElement('ViewCertificate')):
            EasyshellLib.getElement('ButtonYES').Click()
        # loop 60 cycles to wait for RDP session
        time.sleep(2)
        print('等待session窗口')
        if self.wait_element(EasyshellLib.getElement('RDPSessionWindow'), 60):
            print('session窗口出现, 上传测试item')
            # send test logon to remote session
            os.system('echo test_logon>{}'.format(self.local_file))
            ftp = CommonLib.FTPUtils(self.ftp_token['ip'], self.ftp_token['username'], self.ftp_token['password'])
            ftp.change_dir(self.log_path)
            ftp.upload_file(self.local_file, self.local_file)
            if self.remote_file in ftp.get_item_list('.'):
                ftp.delete_file(self.remote_file)
            print('等待vdi返回结果 300秒')
            status = self.get_test_status(ftp, 600)
            ftp.close()
            if not status:
                print('获取vdi返回结果超时(300秒)')
                return False
            print('finished wait get test status', status)
            if status.upper() == 'PASS':
                print('logon pass')
            else:
                print('测试Fail, 收到VDI的结果没有PASS')
            time.sleep(10)
            # check if connection will reconnected with persistent
            if profile['Persistent'] == 'ON':
                time.sleep(profile['Launchdelay'])
                if self.wait_element(
                        EasyshellLib.getElement('Connect', searchFromControl=EasyshellLib.getElement('RDP_WARNNING'))):
                    print('-------------perisitent ON PASS ------------------')
                else:
                    print('-------------perisitent ON Fail ------------------')
                    return False
            else:
                time.sleep(profile['Launchdelay'])
                if self.wait_element(
                        EasyshellLib.getElement('Connect', searchFromControl=EasyshellLib.getElement('RDP_WARNNING'))):
                    print('-------------perisitent OFF Fail ------------------')
                else:
                    print('-------------perisitent OFF PASS ------------------')
                    return False
        else:
            print('logon fail')
            return False


class CitrixLogon(Logon):
    def __init__(self):
        Logon.__init__(self)
        self.session_pane = ''
        self.local_file = 'test_citrix.txt'
        self.remote_file = 'test_citrix_result.txt'
        self.log_path = '/Function/Automation/log/citrix'

    def logon(self, profile):
        if profile['Autolaunch'] == 'OFF' or not profile['Autolaunch']:
            self.utils(profile, 'launch', 'conn')
        time.sleep(profile['Launchdelay'])
        wnd = self.wait_element(EasyshellLib.getElement('CitrixWindows'), 30)
        if wnd:
            print('Citrix启动成功,窗口被打开')
            wnd.Close()
            if self.wait_element(EasyshellLib.getElement('CloseWindowDialog'), 10):
                EasyshellLib.getElement('ButtonOK').Click()


class StoreLogon(Logon):
    def __init__(self):
        Logon.__init__(self)
        self.session_pane = 'displayPanel'
        self.local_file = 'test_storefont.txt'
        self.remote_file = 'test_storefont_result.txt'
        self.log_path = '/Function/Automation/log/storefont'

    def logon(self, profile):
        print('判断是否安装rootca, 如果没装则装到root下面')
        if not self.check_cert('rootca'):
            self.install_ca('.\\test_data\\rootca.cer')
        print('判断autolaunch是否off, 如果为off, 手动启动connection')
        if profile['Autolaunch'] == 'OFF':
            self.utils(profile, 'launch', 'conn')
        time.sleep(profile['Launchdelay'])
        time.sleep(5)
        if not EasyshellLib.getElement('StorePool').Exists():
            print('Storefront桌面池没有启动, 测试FAIL!')
            return False
        if profile['DesktopName']:
            CommonLib.TextControl(Name=profile['DesktopName']).GetParentControl().Click()
            wnd = self.wait_element(
                CommonLib.WindowControl(RegexName='{} - .*'.format(profile['DesktopName'])), 30)
            if wnd:
                print('开始上传test_storefont.txt, 等待15秒')
                os.system('echo test_logon>{}'.format(self.local_file))
                ftp = CommonLib.FTPUtils(self.ftp_token['ip'], self.ftp_token['username'], self.ftp_token['password'])
                ftp.change_dir(self.log_path)
                ftp.upload_file(self.local_file, self.local_file)
                # -------close license expire dialog for windows7- -----------
                time.sleep(15)
                # ------------------------------------------------------------
                print('开始等待vdi的结果返回,300s')
                status = self.get_test_status(ftp, 300)
                ftp.close()
                if not status:
                    print('获取vdi返回结果超时(300秒)')
                    return False
                if status.upper() == 'PASS':
                    print('logon pass')
                    print('检查session窗口{}是否关系(logoff成功)'.format(profile['DesktopName']))
                    if self.wait_element(
                            CommonLib.WindowControl(RegexName='{} - Desktop Viewer'.format(profile['DesktopName'])),
                            180, False):
                        EasyshellLib.getElement('Disconnect').Click()
                    else:
                        print('Logoff 超时!!')
                else:
                    print('测试Fail, 收到VDI的结果没有PASS')
                print('所有Storefont测试结束')
        elif profile['AppName']:
            return True
        else:
            print('No app or desktop need to be launch, logon exit')
            if EasyshellLib.getElement('Disconnect').Exists():
                EasyshellLib.getElement('Disconnect').Click()
            return True


class ViewLogon(Logon):
    """
    MenuBar->MenuItem: if seleted, AccessibleCurrentState()=16
    """

    def __init__(self):
        Logon.__init__(self)
        self.session_pane = 'VMwareSession'
        self.local_file = 'test_vmware.txt'
        self.remote_file = 'test_vmware_result.txt'
        self.log_path = '/Function/Automation/log/vmware'

    def logon(self, profile):
        print('判断是否安装rootca, 如果没装则装到root下面')
        if not self.check_cert('rootca'):
            self.install_ca('.\\test_data\\rootca.cer')
        print('判断autolaunch是否off, 如果为off, 手动启动connection')
        if profile['Autolaunch'] == 'OFF':
            self.utils(profile, 'launch', 'conn')
        time.sleep(profile['Launchdelay'])
        print('检查桌面池...')
        if not self.wait_element(CommonLib.PaneControl(AutomationId='793')):
            print('桌面池没有出现, 测试Fail')
            return False
        time.sleep(3)
        # launch desktop or app
        if profile['AppName']:
            CommonLib.WindowControl(Name='VMware Horizon Client').Click()
            # must add above click, or else will below focus fail
            app_name = CommonLib.PaneControl(Name=profile['AppName'])
            app_name.SetFocus()
            app_name.DoubleClick()
            wait_time = 30000
            current_cycle = 0
            while 1:
                if current_cycle > wait_time:
                    break
                """
                Code here to check app windows show up
                """
        elif profile['DesktopName']:
            print('让需要启动的桌面图标{}显示出来'.format(profile['DesktopName']))
            CommonLib.WindowControl(Name='VMware Horizon Client').Click()
            # must add above click, or else will below focus fail
            desktop_name = CommonLib.PaneControl(Name=profile['DesktopName'])
            desktop_name.SetFocus()
            desktop_name.DoubleClick()
            print('检查VMWare Session窗口{}是否出现'.format(profile['DesktopName']))
            if self.wait_element(CommonLib.WindowControl(Name=profile['DesktopName']), 30):
                print('开始上传test_vmware.txt, 等待15秒')
                os.system('echo test_logon>{}'.format(self.local_file))
                ftp = CommonLib.FTPUtils(self.ftp_token['ip'], self.ftp_token['username'], self.ftp_token['password'])
                ftp.change_dir(self.log_path)
                ftp.upload_file(self.local_file, self.local_file)
                # -------close license expire dialog for windows7- -----------
                time.sleep(15)
                print('开始点击license对话框')
                x, y = self.get_resolution()
                CommonLib.Click(int(x / 2), int(y / 3))
                CommonLib.SendKey(CommonLib.Keys.VK_ESCAPE)
                # ------------------------------------------------------------
                print('开始等待vdi的结果返回,300s')
                status = self.get_test_status(ftp, 300)
                ftp.close()
                if not status:
                    print('获取vdi返回结果超时(300秒)')
                    return False
                print('finished wait get test status', status)
                if status.upper() == 'PASS':
                    print('logon pass')
                    print('检查session窗口{}是否关系(logoff成功)'.format(profile['DesktopName']))
                    if self.wait_element(CommonLib.WindowControl(Name=profile['DesktopName']), 180, False):
                        EasyshellLib.getElement('VMwarePool').Close()
                    else:
                        print('Logoff 超时!!')
                else:
                    print('测试Fail, 收到VDI的结果没有PASS')
                print('所有VMWare测试结束')

        else:
            print('do not need to launch session or app')
            return
        # wait 10 min for app or session launched


if __name__ == '__main__':
    # rdp_profile = dict(
    #     Name='test_rdp',
    #     Password='Shanghai2010',
    #     Username='Administrator',
    #     Hostname='15.83.248.204',
    #     Autolaunch="OFF",
    #     Launchdelay=0,
    #     Persistent='OFF'
    # )
    view_profile = dict(
        Name='test_view',
        Password='zhao123',
        Username='zhao.sam',
        Domain='sh.dto',
        Hostname='vnsvr.sh.dto',
        Autolaunch="OFF",
        Launchdelay=0,
        Persistent='OFF',
        Layout='fullscreen',
        ConnUSBStartup='OFF',
        ConnUSBInsertion='OFF',
        DesktopName='Windows7C',
        AppName=None
    )
    # citrix_profile = dict(
    #     Name='test_citrix',
    #     Password='zhao123',
    #     Username='zhao.sam',
    #     Domain='sh.dto',
    #     Hostname='XA65FP2SVR03.SH.DTO',
    #     Autolaunch="OFF",
    #     Launchdelay=0,
    #     Persistent='OFF',
    #     Remotesize='fullscreen',
    # )
    # storefont_profile = dict(
    #     Name='test_storefont',
    #     Password='zhao123',
    #     Username='zhao.sam',
    #     Domain='sh.dto',
    #     Hostname='fcxds.sh.dto',
    #     Autolaunch="OFF",
    #     Launchdelay=0,
    #     Persistent='OFF',
    #     CustomLogon=None,
    #     DesktopToolbar="OFF",
    #     DesktopName='Win10',
    #     AppName=None
    # )
    # CitrixLogon().logon(citrix_profile)
    RDPLogon().logoff()
    # StoreLogon().logon(storefont_profile)
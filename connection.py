import shutil
import Test_Scripts.EasyShell_Lib as EasyshellLib
import Library.CommonLib as CommonLib
import os
import time
import traceback
import pysnooper

"""
interaction command between local and remote session:
local file: test_vdi.txt (ex. test_rdp.txt)
remote file: test_vdi_result.txt (ex. test_rdp_result.txt)
test logon: send by local, remote get this command should logoff session
test xxxxx: send by local, remote get this command should test function xxxxx
test status: send by local, remote get this command should return vdi's test status "test_status_Finished | test_status_testing | None"
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
        return flag.strip()

    @staticmethod
    def wait_element(element, cycles=3):
        flag = None
        for i in range(cycles):
            if element.Exists(0, 0):
                flag = element
                break
            else:
                flag = None
                print('loop', i)
                time.sleep(1)
                continue
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
        return


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
        if self.wait_element(
                EasyshellLib.getElement('Connect', searchFromControl=EasyshellLib.getElement('RDP_WARNNING'))):
            EasyshellLib.getElement('Connect',
                                    searchFromControl=EasyshellLib.getElement('RDP_WARNNING')).Click()
        else:
            print('-----------Manual Launch fail, connection not launch------------------------')
            return False
        # loop 3 times to wait for warning dialog pops up
        time.sleep(1)
        if self.wait_element(EasyshellLib.getElement('RDPPassword')):
            EasyshellLib.getElement('RDPPassword').SetValue(profile['Password'])
            EasyshellLib.getElement('ButtonOK').Click()
        # loop 3 times wait for certificate dialog pops up
        time.sleep(1)
        if self.wait_element(EasyshellLib.getElement('ViewCertificate')):
            EasyshellLib.getElement('ButtonYES').Click()
        # loop 60 cycles to wait for RDP session
        time.sleep(1)
        if self.wait_element(EasyshellLib.getElement('RDPSessionWindow'), 60):
            # send test logon to remote session
            os.system('echo test_logon>{}'.format(self.local_file))
            ftp = CommonLib.FTPUtils(self.ftp_token['ip'], self.ftp_token['username'], self.ftp_token['password'])
            ftp.change_dir(self.log_path)
            ftp.upload_file(self.local_file, self.local_file)
            status = self.get_test_status(ftp, 300)
            print('finished wait get test status', status)
            if status.upper() == 'PASS':
                print('logon pass')
            else:
                print('logon fail, no result return, maybe not logoff')
            ftp.close()
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
    pass


class StoreLogon(Logon):
    pass


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

    @staticmethod
    def install_ca(file):
        EasyshellLib.CommonUtils.import_cert(file)

    @staticmethod
    def check_cert(name):
        return EasyshellLib.CommonUtils.get_cert_id(name)

    @staticmethod
    def del_cert(index):
        EasyshellLib.CommonUtils.del_cert(index)

    def logon(self, profile):
        if not self.check_cert('rootca'):
            self.install_ca('.\\test_data\\rootca.cer')
        if profile['Autolaunch'] == 'OFF' or not profile['Autolaunch']:
            self.utils(profile, 'launch', 'conn')
        time.sleep(profile['Launchdelay'])
        if not self.wait_element(CommonLib.PaneControl(AutomationId='793')):
            print('Desktop pool not shown, test fail')
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
                if current_cycle>wait_time:
                    break
                """
                Code here to check app windows show up
                """
        elif profile['DesktopName']:
            CommonLib.WindowControl(Name='VMware Horizon Client').Click()
            # must add above click, or else will below focus fail
            desktop_name = CommonLib.PaneControl(Name=profile['DesktopName'])
            desktop_name.SetFocus()
            desktop_name.DoubleClick()
            wnd = Logon().wait_element(CommonLib.WindowControl(Name=profile['DesktopName']), 5)
            if wnd:
                print('Find wnd')
                os.system('echo test_logon>{}'.format(self.local_file))
                ftp = CommonLib.FTPUtils(self.ftp_token['ip'], self.ftp_token['username'], self.ftp_token['password'])
                ftp.change_dir(self.log_path)
                ftp.upload_file(self.local_file, self.local_file)
                # -------close license expire dialog for windows7- -----------
                time.sleep(15)
                x, y = self.get_resolution()
                CommonLib.Click(int(x/2), int(y/3))
                CommonLib.SendKey(CommonLib.Keys.VK_ESCAPE)
                # ------------------------------------------------------------
                status = self.get_test_status(ftp, 300)
                print('finished wait get test status', status)
                if status.upper() == 'PASS':
                    print('logon pass')
                else:
                    print('logon fail, no result return, maybe not logoff')
                ftp.close()

        else:
            print('do not need to launch session or app')
            return
        # wait 10 min for app or session launched


class ViewRDP(ViewLogon):
    pass


class ViewPCOIP(ViewLogon):
    pass


class ViewBlast(ViewLogon):
    pass


class VDITest:
    def __init__(self):
        self.test_req = None
        self.local_file = ''
        self.remote_file = ''
        self.log_path = ''
        self.log_path_list = ['/Function/Automation/log/rdp', '/Function/Automation/log/vmware',
                         '/Function/Automation/log/storefont', '/Function/Automation/log/citrix']
        self.tool_path = '/Function/Automation/vditest'

    def get_test(self):
        ftp = CommonLib.FTPUtils('15.83.251.201', 'sh\\cheng.balance', 'password.321')
        for log_path in self.log_path_list:
            vdi_item = os.path.basename(log_path)
            self.local_file = 'test_{}.txt'.format(vdi_item)
            self.remote_file = 'test_{}_result.txt'.format(vdi_item)
            ftp.change_dir(log_path)
            if self.local_file in ftp.get_item_list(log_path):
                self.log_path = log_path
                ftp.download_file(self.local_file, self.local_file)
                with open(self.local_file, 'r') as f:
                    _, test_item = f.read().strip().split('_')
                ftp.close()
                return vdi_item, test_item
        if self.log_path == '':
                print('No test needed')


    def download_test(self):
        self.test_req = self.get_test()
        if not os.path.exists('c:\\scripts'):
            os.mkdir('c:\\scripts')
        if self.test_req:
            print(self.test_req[1].strip())
            if os.path.exists('c:\\scripts\\{}'.format(self.test_req[1].strip())):
                shutil.rmtree('c:\\scripts\\{}'.format(self.test_req[1].strip()))
            ftp = CommonLib.FTPUtils('15.83.251.201', 'sh\\cheng.balance', 'password.321')
            ftp.change_dir(self.tool_path + '/{}'.format(self.test_req[0]))
            print(ftp.get_item_list('.'))
            ftp.download_dir(self.test_req[1].strip(), 'c:\\scripts\\{}'.format(self.test_req[1].strip()))
            ftp.change_dir(self.log_path)
            ftp.delete_file(self.local_file)
            ftp.close()
            return True
        else:
            print('Fail, no test requirement upload to ftp')
            return False

    def run_test(self):
        print(self.test_req, '---------------------------')
        if self.test_req:
            os.system(r'c:\scripts\{}\start.exe'.format(self.test_req[1]))

    def get_status(self):
        flag = None
        if not os.path.exists('c:\\scripts\\{}\\flag.txt'.format(self.test_req[1])):
            flag = None
        with open('c:\\scripts\\{}\\flag.txt'.format(self.test_req[1])) as f:
            status = f.read().strip()
            if 'PASS' in status.upper():
                flag = 'PASS'
            elif 'RUNNING' in status.upper():
                flag = 'RUNNING'
            else:
                flag = status.split(' ')[-1]
        return flag

    def start(self):
        while 1:
            if self.download_test():
                self.run_test()
                self.test_req = None
            else:
                print('wait')
                time.sleep(5)
                continue


if __name__ == '__main__':
    VDITest().start()
    # rdp_profile = dict(
    #     Name='test_rdp',
    #     Password='Shanghai2010',
    #     Username='Administrator',
    #     Hostname='15.83.248.204',
    #     Autolaunch="OFF",
    #     Launchdelay=0,
    #     Persistent='OFF'
    # )
    # view_profile = dict(
    #     Name='test_view',
    #     Password='zhao123',
    #     Username='zhao.sam',
    #     Domain='sh.dto',
    #     Hostname='vnsvr.sh.dto',
    #     Autolaunch="OFF",
    #     Launchdelay=0,
    #     Persistent='OFF',
    #     Layout='fullscreen',
    #     ConnUSBStartup='OFF',
    #     ConnUSBInsertion='OFF',
    #     DesktopName='Windows7C',
    #     AppName=None
    # )
    # ViewLogon().logon(view_profile)
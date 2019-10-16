import os
import shutil
import time
from Library import CommonLib


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